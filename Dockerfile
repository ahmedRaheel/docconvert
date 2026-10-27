FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 DEBIAN_FRONTEND=noninteractive

RUN apt-get -o Acquire::Retries=5 update \
 && apt-get -o Acquire::Retries=5 install -y --no-install-recommends \
      fonts-liberation fonts-dejavu-core libpango-1.0-0 libpangoft2-1.0-0 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY pyproject.toml README.md ./
COPY docconvert ./docconvert
COPY requirements-web.txt ./
RUN pip install ".[fidelity,html]" -r requirements-web.txt

COPY web ./web
RUN useradd --create-home --uid 1001 app && mkdir /data && chown app:app /data
USER app
ENV WORK_DIR=/data HOME=/home/app

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s \
  CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/healthz',timeout=4).status==200 else 1)"

CMD ["gunicorn", "--workers", "1", "--threads", "2", "--timeout", "300", "--bind", "0.0.0.0:8000", "--access-logfile", "-", "web.app:app"]