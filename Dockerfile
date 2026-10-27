FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 DEBIAN_FRONTEND=noninteractive

# LibreOffice = highest-fidelity Word/Excel/PowerPoint -> PDF.
# Carlito/Caladea are metric-compatible with Calibri/Cambria, so Office layouts keep their line breaks.
# Pango/HarfBuzz are needed by WeasyPrint (HTML -> PDF with full CSS).
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-writer libreoffice-calc libreoffice-impress \
      fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea fonts-noto-core \
      libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0 \
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

# Long timeout: big Office files can take a while in LibreOffice.
CMD ["gunicorn", "--workers", "2", "--threads", "4", "--timeout", "300", "--bind", "0.0.0.0:8000", "--access-logfile", "-", "web.app:app"]
