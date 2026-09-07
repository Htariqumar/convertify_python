FROM python:3.11-slim

# LibreOffice (word/excel/ppt <-> pdf, office thumbnails) and Ghostscript (compress-pdf,
# pdf thumbnails) - both installed as normal Debian packages since this runs on a regular
# Linux host, not a size-capped serverless function. fonts-* packages keep converted
# documents from silently substituting missing glyphs.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    python3-uno \
    python3-pip \
    ghostscript \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# main.py keeps one LibreOffice instance running persistently (via `unoserver`, see
# USE_PERSISTENT_LIBREOFFICE in main.py) instead of spawning a fresh one per request - that
# listener process needs LibreOffice's `uno` Python bridge, which is only importable from
# Debian's system python3 (just installed via python3-uno above), not this image's own
# /usr/local/bin/python3 that runs the FastAPI app itself (see `pip install` below, and the
# `unoserver` *client* class the app uses instead, which needs no uno import of its own).
# --break-system-packages because Debian's system pip otherwise refuses a global install.
RUN /usr/bin/python3 -m pip install --no-cache-dir --break-system-packages unoserver

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8001
# Each worker is a separate process that loads its own copy of the Whisper model into
# RAM, so raising this multiplies memory use accordingly - only increase it once the
# host has RAM to spare (see DEPLOYMENT.md). Default of 1 is safe for small hosts and
# still handles concurrent requests fine (see run_in_threadpool usage in main.py) -
# this just adds true multi-process parallelism for when traffic grows.
ENV UVICORN_WORKERS=1
# Tells main.py's persistent-LibreOffice-listener feature which Python has the `uno` bridge
# installed above - see the comment on UNOSERVER_PYTHON in main.py.
ENV UNOSERVER_PYTHON=/usr/bin/python3
EXPOSE 8001

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT} --workers ${UVICORN_WORKERS}"]
