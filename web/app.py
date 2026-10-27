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
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file
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


threading.Thread(target=_janitor, daemon=True, name="janitor").start()


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
        "font-src https://fonts.gstatic.com https://cdn.jsdelivr.net; img-src 'self' data:; connect-src 'self' https://cdn.jsdelivr.net")
    if request.is_secure:  # only behind HTTPS (the reverse proxy sets X-Forwarded-Proto)
        resp.headers.setdefault("Strict-Transport-Security", "max-age=31536000")
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


def _err(message: str, status: int):
    return jsonify(error=message), status


@app.errorhandler(413)
def _too_big(_):
    return _err(f"That file is larger than {MAX_MB} MB.", 413)


# ----------------------------------------------------------------- pages & api
@app.get("/")
def index():
    boot = {"base": request.script_root, "maxMb": MAX_MB, "formats": FORMATS, "canon": CANON}
    return render_template("index.html", app_name=APP_NAME, max_mb=MAX_MB, ttl_minutes=max(1, TTL // 60), boot=boot)


@app.get("/healthz")
def health():
    return {"status": "ok"}


@app.post("/api/convert")
def api_convert():
    f = request.files.get("file")
    to = (request.form.get("to") or "").lower()
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
        raw = raw if isinstance(raw, dict) else {}
    except ValueError:
        raw = {}
    opts: dict = {}
    if raw.get("pages"):
        if not PAGES.match(str(raw["pages"])):
            return _err("Pages should look like 1-3, 5.", 400)
        opts["pages"] = str(raw["pages"]).replace(" ", "")
    if raw.get("mode") in ("editable", "image"):
        opts["mode"] = raw["mode"]
    if raw.get("sheet_mode") in ("page", "table"):
        opts["sheet_mode"] = raw["sheet_mode"]
    engine = raw.get("engine") if raw.get("engine") in ("auto", "builtin") else "auto"

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


@app.get("/api/download/<job>")
def download(job):
    f = _result(job)
    if not f:
        return _err("This file has expired. Convert it again.", 404)
    return send_file(f, as_attachment=True, download_name=f.name, max_age=0)


@app.get("/api/download-all")
def download_all():
    ids = [i for i in (request.args.get("ids") or "").split(",") if i][:25]
    files = [f for f in (_result(i) for i in ids) if f]
    if not files:
        return _err("These files have expired. Convert them again.", 404)
    buf, seen = io.BytesIO(), {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            n = seen.get(f.name, 0)
            seen[f.name] = n + 1
            z.write(f, f.name if not n else f"{f.stem} ({n}){f.suffix}")
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"{APP_NAME.lower()}-converted.zip", max_age=0)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=False)
