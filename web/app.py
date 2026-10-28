"""Pagewise web interface for docconvert (Flask).

Stateless by design: every result lives in ``WORK_DIR/<job-id>/out`` so any
gunicorn worker can serve a download, and a janitor thread removes old jobs.
"""
from __future__ import annotations

import io
import json
import logging
import os
import re
import secrets
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, render_template, request, send_file
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

import docconvert
from docconvert import DocConvertError, MissingDependencyError

APP_NAME = os.getenv("APP_NAME", "Pagewise")
MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))
TTL = int(os.getenv("RESULT_TTL_SECONDS", "1800"))
MAX_PARALLEL = int(os.getenv("MAX_CONCURRENT_CONVERSIONS", "2"))
TIMEOUT = int(os.getenv("CONVERSION_TIMEOUT", "180"))
WORK = Path(os.getenv("WORK_DIR", tempfile.gettempdir())) / "pagewise"
WORK.mkdir(parents=True, exist_ok=True)

# S3 mode: browsers upload straight to the bucket and results are downloaded straight from it,
# so the EC2 box never proxies file bytes. Enabled by setting S3_BUCKET.
S3_BUCKET = os.getenv("S3_BUCKET", "").strip()
S3_REGION = os.getenv("AWS_REGION") or os.getenv("AWS_DEFAULT_REGION") or None
S3_PREFIX = os.getenv("S3_PREFIX", "pagewise").strip("/")
USE_S3 = bool(S3_BUCKET)

# what each source type can be turned into (legacy types are converted via LibreOffice)
CANON = {"pdf": "pdf", "docx": "docx", "xlsx": "xlsx", "pptx": "pptx", "html": "html",
         "doc": "docx", "odt": "docx", "rtf": "docx", "xls": "xlsx", "ods": "xlsx", "ppt": "pptx", "odp": "pptx"}
BASE_TARGETS = {"pdf": ["docx", "xlsx", "pptx"], "docx": ["pdf", "xlsx", "pptx"], "xlsx": ["pdf", "docx", "pptx"],
                "pptx": ["pdf", "docx", "xlsx"], "html": ["pdf", "docx", "xlsx", "pptx"]}
FORMATS = {ext: BASE_TARGETS[c] for ext, c in CANON.items()}
if not docconvert.available_engines().get("libreoffice"):
    FORMATS = {k: v for k, v in FORMATS.items() if k == CANON[k]}  # legacy types need LibreOffice

JOB_ID = re.compile(r"^[0-9a-f]{32}$")
PAGES = re.compile(r"^[0-9,\-\s]{1,64}$")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
log = logging.getLogger("pagewise")
_slots = threading.BoundedSemaphore(MAX_PARALLEL)
_pool = ThreadPoolExecutor(max_workers=MAX_PARALLEL, thread_name_prefix="convert")
_s3 = None


def s3():
    """Lazily created S3 client (credentials come from the EC2 instance role)."""
    global _s3
    if _s3 is None:
        import boto3
        from botocore.config import Config
        _s3 = boto3.client("s3", region_name=S3_REGION, config=Config(
            signature_version="s3v4", s3={"addressing_style": "virtual"},
            retries={"max_attempts": 4, "mode": "standard"}))
    return _s3


def _key(*parts: str) -> str:
    return "/".join(([S3_PREFIX] if S3_PREFIX else []) + list(parts))


def _put_status(job: str, data: dict):
    s3().put_object(Bucket=S3_BUCKET, Key=_key("jobs", job + ".json"), Body=json.dumps(data).encode(),
                    ContentType="application/json")


def _get_status(job: str):
    if not JOB_ID.match(job or ""):
        return None
    try:
        body = s3().get_object(Bucket=S3_BUCKET, Key=_key("jobs", job + ".json"))["Body"].read()
        return json.loads(body)
    except Exception:
        return None


# ----------------------------------------------------------------- housekeeping
def _janitor():
    while True:
        time.sleep(120)
        now = time.time()
        try:
            for d in WORK.iterdir():
                if d.is_dir() and now - d.stat().st_mtime > TTL:
                    shutil.rmtree(d, ignore_errors=True)
        except Exception:  # never let the janitor die
            log.exception("cleanup failed")


threading.Thread(target=_janitor, daemon=True, name="janitor").start()  # S3 mode uses a bucket lifecycle rule


@app.before_request
def _guard():
    user, pw = os.getenv("BASIC_AUTH_USER"), os.getenv("BASIC_AUTH_PASSWORD")
    if user and pw and request.endpoint != "health":
        a = request.authorization
        ok = a and secrets.compare_digest(a.username or "", user) and secrets.compare_digest(a.password or "", pw)
        if not ok:
            return Response("Sign in to use this tool.", 401, {"WWW-Authenticate": f'Basic realm="{APP_NAME}"'})


@app.after_request
def _headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("Referrer-Policy", "same-origin")
    resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    resp.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; script-src 'self' https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com https://cdn.jsdelivr.net; img-src 'self' data:; connect-src 'self' https://*.amazonaws.com")
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


def _err(message: str, status: int):
    return jsonify(error=message), status


@app.errorhandler(413)
def _too_big(_):
    return _err(f"That file is larger than {MAX_MB} MB.", 413)



def _clean_options(raw):
    """Whitelist what the browser may pass on to the converter -> (opts, engine, error)."""
    raw = raw if isinstance(raw, dict) else {}
    opts: dict = {}
    if raw.get("pages"):
        if not PAGES.match(str(raw["pages"])):
            return {}, "auto", "Pages should look like 1-3, 5."
        opts["pages"] = str(raw["pages"]).replace(" ", "")
    if raw.get("mode") in ("editable", "image"):
        opts["mode"] = raw["mode"]
    if raw.get("layout") in ("auto", "fixed", "flow"):
        opts["layout"] = raw["layout"]
    if raw.get("sheet_mode") in ("page", "table"):
        opts["sheet_mode"] = raw["sheet_mode"]
    engine = raw.get("engine") if raw.get("engine") in ("auto", "builtin") else "auto"
    return opts, engine, None


def _explain(exc: Exception):
    """Exception -> (message, http status)."""
    if isinstance(exc, MissingDependencyError):
        return f"This conversion needs a component that isn't installed on the server: {exc}", 501
    if isinstance(exc, ValueError):
        return str(exc), 400
    if isinstance(exc, DocConvertError):
        return f"The file couldn't be converted: {exc}", 422
    log.exception("conversion crashed", exc_info=exc)
    return "Something went wrong while converting. Try again, or try the built-in engine under More options.", 500


# ----------------------------------------------------------------- pages & api
@app.get("/")
def index():
    boot = {"base": request.script_root, "maxMb": MAX_MB, "formats": FORMATS, "canon": CANON,
            "storage": "s3" if USE_S3 else "local"}
    return render_template("index.html", app_name=APP_NAME, max_mb=MAX_MB, ttl_minutes=max(1, TTL // 60), boot=boot)


@app.get("/healthz")
def health():
    return {"status": "ok", "storage": "s3" if USE_S3 else "local"}


@app.post("/api/convert")
def api_convert():
    f = request.files.get("file")
    to = (request.form.get("to") or "").lower()
    if USE_S3:
        return _err("Upload the file to storage first (see /api/uploads).", 400)
    if not f or not f.filename:
        return _err("Choose a file to convert.", 400)
    name = secure_filename(f.filename) or "document"
    ext = Path(name).suffix.lower().lstrip(".")
    ext = "html" if ext == "htm" else ext
    if ext not in FORMATS:
        return _err("That file type isn't supported. Use PDF, Word, Excel, PowerPoint or HTML.", 415)
    if to not in FORMATS[ext]:
        return _err(f"A .{ext} file can't be converted to .{to}.", 400)

    try:
        raw = json.loads(request.form.get("options") or "{}")
    except ValueError:
        raw = {}
    opts, engine, problem = _clean_options(raw)
    if problem:
        return _err(problem, 400)

    job = uuid.uuid4().hex
    jdir = WORK / job
    (jdir / "in").mkdir(parents=True)
    (jdir / "out").mkdir()
    src = jdir / "in" / f"source.{ext}"
    out = jdir / "out" / (Path(name).stem[:120] + "." + to)
    try:
        f.save(src)
        if not _slots.acquire(timeout=120):
            return _err("The converter is busy. Try again in a minute.", 503)
        try:
            docconvert.convert(src, out, to=to, engine=engine, timeout=TIMEOUT, **opts)
        finally:
            _slots.release()
    except MissingDependencyError as exc:
        shutil.rmtree(jdir, ignore_errors=True)
        return _err(f"This conversion needs a component that isn't installed on the server: {exc}", 501)
    except ValueError as exc:
        shutil.rmtree(jdir, ignore_errors=True)
        return _err(str(exc), 400)
    except DocConvertError as exc:
        shutil.rmtree(jdir, ignore_errors=True)
        return _err(f"The file couldn't be converted: {exc}", 422)
    except Exception:
        log.exception("conversion crashed")
        shutil.rmtree(jdir, ignore_errors=True)
        return _err("Something went wrong while converting. Try again, or try the built-in engine under More options.", 500)
    finally:
        shutil.rmtree(jdir / "in", ignore_errors=True)

    return jsonify(id=job, name=out.name, size=out.stat().st_size, to=to)


def _result(job: str) -> Path | None:
    if not JOB_ID.match(job or ""):
        return None
    files = list((WORK / job / "out").glob("*"))
    return files[0] if files else None


def _source_ext(name: str):
    ext = Path(name).suffix.lower().lstrip(".")
    return "html" if ext == "htm" else ext


@app.post("/api/uploads")
def api_uploads():
    """Hand the browser a short-lived, size-limited form to upload one file straight to S3."""
    if not USE_S3:
        return _err("Direct upload isn't enabled on this server.", 404)
    data = request.get_json(silent=True) or {}
    name = secure_filename(str(data.get("name") or "")) or "document"
    ext = _source_ext(name)
    if ext not in FORMATS:
        return _err("That file type isn't supported. Use PDF, Word, Excel, PowerPoint or HTML.", 415)
    try:
        size = int(data.get("size") or 0)
    except (TypeError, ValueError):
        size = 0
    if size > MAX_MB * 1024 * 1024:
        return _err(f"That file is larger than {MAX_MB} MB.", 413)
    key = _key("uploads", uuid.uuid4().hex, name[:150])
    post = s3().generate_presigned_post(S3_BUCKET, key, ExpiresIn=900,
                                        Conditions=[["content-length-range", 1, MAX_MB * 1024 * 1024]])
    return jsonify(url=post["url"], fields=post["fields"], key=key)


def _run_s3_job(job, key, ext, to, opts, engine, out_name):
    """Runs in a worker thread: S3 -> temp file -> convert -> S3, recording progress in S3."""
    try:
        _put_status(job, {"state": "working"})
        with tempfile.TemporaryDirectory(dir=WORK) as td:
            src, out = Path(td) / f"source.{ext}", Path(td) / out_name
            s3().download_file(S3_BUCKET, key, str(src))
            docconvert.convert(src, out, to=to, engine=engine, timeout=TIMEOUT, **opts)
            rkey = _key("results", job, out_name)
            s3().upload_file(str(out), S3_BUCKET, rkey)
            _put_status(job, {"state": "done", "name": out_name, "size": out.stat().st_size, "key": rkey})
    except Exception as exc:
        msg, _ = _explain(exc)
        try:
            _put_status(job, {"state": "failed", "error": msg})
        except Exception:
            log.exception("could not record failure")
    finally:
        try:
            s3().delete_object(Bucket=S3_BUCKET, Key=key)   # never keep the original longer than needed
        except Exception:
            log.warning("could not delete %s", key)


@app.post("/api/s3/convert")
def api_s3_convert():
    if not USE_S3:
        return _err("Direct upload isn't enabled on this server.", 404)
    data = request.get_json(silent=True) or {}
    key, to = str(data.get("key") or ""), str(data.get("to") or "").lower()
    pattern = re.compile("^" + re.escape(_key("uploads")) + r"/[0-9a-f]{32}/[A-Za-z0-9._\-]{1,150}$")
    if not pattern.match(key):
        return _err("That upload wasn't recognised.", 400)
    name = key.rsplit("/", 1)[1]
    ext = _source_ext(name)
    if ext not in FORMATS or to not in FORMATS[ext]:
        return _err(f"A .{ext} file can't be converted to .{to}.", 400)
    opts, engine, problem = _clean_options(data.get("options"))
    if problem:
        return _err(problem, 400)
    try:
        head = s3().head_object(Bucket=S3_BUCKET, Key=key)
    except Exception:
        return _err("The upload didn't arrive. Try again.", 404)
    if head["ContentLength"] > MAX_MB * 1024 * 1024:
        s3().delete_object(Bucket=S3_BUCKET, Key=key)
        return _err(f"That file is larger than {MAX_MB} MB.", 413)
    job = uuid.uuid4().hex
    _put_status(job, {"state": "queued"})
    _pool.submit(_run_s3_job, job, key, ext, to, opts, engine, Path(name).stem[:120] + "." + to)
    return jsonify(job=job), 202


@app.get("/api/jobs/<job>")
def api_job(job):
    st = _get_status(job) if USE_S3 else None
    if not st:
        return _err("Unknown or expired job.", 404)
    return jsonify({k: st[k] for k in ("state", "name", "size", "error") if k in st})


@app.get("/api/download/<job>")
def download(job):
    if USE_S3:
        st = _get_status(job)
        if not st or st.get("state") != "done":
            return _err("This file has expired. Convert it again.", 404)
        url = s3().generate_presigned_url(
            "get_object", ExpiresIn=300,
            Params={"Bucket": S3_BUCKET, "Key": st["key"],
                    "ResponseContentDisposition": f'attachment; filename="{st["name"]}"'})
        return redirect(url, 302)
    f = _result(job)
    if not f:
        return _err("This file has expired. Convert it again.", 404)
    return send_file(f, as_attachment=True, download_name=f.name, max_age=0)


@app.get("/api/download-all")
def download_all():
    ids = [i for i in (request.args.get("ids") or "").split(",") if i][:25]
    tmp = tempfile.TemporaryDirectory(dir=WORK)
    if USE_S3:
        files = []
        for i in ids:
            st = _get_status(i)
            if st and st.get("state") == "done":
                dest = Path(tmp.name) / i / st["name"]
                dest.parent.mkdir()
                s3().download_file(S3_BUCKET, st["key"], str(dest))
                files.append(dest)
    else:
        files = [f for f in (_result(i) for i in ids) if f]
    if not files:
        tmp.cleanup()
        return _err("These files have expired. Convert them again.", 404)
    buf, seen = io.BytesIO(), {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            n = seen.get(f.name, 0)
            seen[f.name] = n + 1
            z.write(f, f.name if not n else f"{f.stem} ({n}){f.suffix}")
    buf.seek(0)
    tmp.cleanup()
    return send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"{APP_NAME.lower()}-converted.zip", max_age=0)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=False)
