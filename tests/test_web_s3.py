import io
import time

import pytest

pytest.importorskip("flask")
boto3 = pytest.importorskip("boto3")
moto = pytest.importorskip("moto")

import web.app as webapp  # noqa: E402

BUCKET = "pagewise-test"


@pytest.fixture()
def s3env(monkeypatch):
    with moto.mock_aws():
        monkeypatch.setattr(webapp, "USE_S3", True)
        monkeypatch.setattr(webapp, "S3_BUCKET", BUCKET)
        monkeypatch.setattr(webapp, "S3_REGION", "us-east-1")
        monkeypatch.setattr(webapp, "_s3", None)
        monkeypatch.setenv("AWS_ACCESS_KEY_ID", "x"); monkeypatch.setenv("AWS_SECRET_ACCESS_KEY", "x")
        boto3.client("s3", region_name="us-east-1").create_bucket(Bucket=BUCKET)
        webapp.app.config["TESTING"] = True
        yield webapp.app.test_client()


def _docx():
    from docx import Document
    d = Document(); d.add_paragraph("hello s3"); b = io.BytesIO(); d.save(b); return b.getvalue()


def test_s3_roundtrip(s3env):
    c = s3env
    assert c.get("/healthz").json["storage"] == "s3"
    r = c.post("/api/uploads", json={"name": "My Doc.docx", "size": 1000})
    assert r.status_code == 200
    meta = r.json
    assert "content-length-range" in str(meta["fields"]) or meta["fields"]["key"] == meta["key"]
    # simulate the browser's direct upload
    webapp.s3().put_object(Bucket=BUCKET, Key=meta["key"], Body=_docx())

    r = c.post("/api/s3/convert", json={"key": meta["key"], "to": "pdf", "options": {"engine": "builtin"}})
    assert r.status_code == 202, r.json
    job = r.json["job"]
    for _ in range(60):
        st = c.get(f"/api/jobs/{job}").json
        if st["state"] in ("done", "failed"):
            break
        time.sleep(0.5)
    assert st["state"] == "done", st
    assert st["name"] == "My_Doc.pdf" and "key" not in st

    d = c.get(f"/api/download/{job}")
    assert d.status_code == 302 and "amazonaws.com" in d.headers["Location"] and "attachment" in d.headers["Location"]
    z = c.get(f"/api/download-all?ids={job}")
    assert z.status_code == 200 and z.data[:2] == b"PK"
    # original upload is removed after conversion
    keys = [o["Key"] for o in webapp.s3().list_objects_v2(Bucket=BUCKET).get("Contents", [])]
    assert not any("/uploads/" in k for k in keys)


@pytest.mark.parametrize("payload,status", [
    ({"name": "a.exe", "size": 10}, 415),
    ({"name": "a.pdf", "size": 10**10}, 413),
])
def test_upload_rejections(s3env, payload, status):
    assert s3env.post("/api/uploads", json=payload).status_code == status


def test_convert_rejects_foreign_keys(s3env):
    for key in ("../../etc/passwd", "pagewise/results/x/y.pdf", "other/uploads/" + "0" * 32 + "/a.pdf"):
        assert s3env.post("/api/s3/convert", json={"key": key, "to": "docx"}).status_code == 400
    assert s3env.post("/api/s3/convert", json={"key": f"pagewise/uploads/{'0'*32}/a.pdf", "to": "docx"}).status_code == 404


def test_conversion_error_is_reported(s3env):
    c = s3env
    key = c.post("/api/uploads", json={"name": "bad.pdf", "size": 9}).json["key"]
    webapp.s3().put_object(Bucket=BUCKET, Key=key, Body=b"not a pdf")
    job = c.post("/api/s3/convert", json={"key": key, "to": "docx"}).json["job"]
    for _ in range(60):
        st = c.get(f"/api/jobs/{job}").json
        if st["state"] in ("done", "failed"):
            break
        time.sleep(0.5)
    assert st["state"] == "failed" and st["error"]
