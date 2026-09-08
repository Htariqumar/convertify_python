FROM python:3.11-slim

# LibreOffice (word/excel/ppt <-> pdf, office thumbnails) and Ghostscript (compress-pdf,
# pdf thumbnails) - both installed as normal Debian packages since this runs on a regular
# Linux host, not a size-capped serverless function. fonts-* packages keep converted
# documents from silently substituting missing glyphs.
#
# fonts-liberation covers Arial/Times New Roman/Courier New (metric-compatible), but most
# .docx files created in modern Word default to Calibri (body) and Cambria (headings) -
# with neither font installed, LibreOffice's DOCX import silently substitutes a font with
# different character widths, which reflows line breaks/spacing/pagination differently
# than Word itself would lay the same document out. fonts-crosextra-carlito/-caladea are
# Google's metric-compatible substitutes for exactly those two fonts (same glyph widths,
# so text wraps at the same points Word would use), and are the single highest-impact fix
# for docx/pptx -> pdf formatting drift from the source document.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    python3-uno \
    python3-pip \
    ghostscript \
    fontconfig \
    fonts-dejavu \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-noto-color-emoji \
    && rm -rf /var/lib/apt/lists/*

# Word's newest default theme font, "Aptos" (2023, replacing Calibri), has no freely
# redistributable metric-compatible clone the way Carlito/Caladea cover Calibri/Cambria -
# Microsoft's own Aptos font files require a separate license to install on a server (see
# fontconfig-aliases.conf for the full explanation). This aliases "Aptos"/"Aptos Display"
# to the Carlito/Caladea already installed above, so LibreOffice gets a reasonably close,
# consistent substitute instead of fontconfig's own generic (and worse-matching) fallback.
# It also aliases "Segoe UI Emoji" (the font Word tags emoji/symbol characters with, even in
# plain text documents) to fonts-noto-color-emoji above - without it, none of the fonts here
# have those glyphs at all, so emoji/dingbat characters (checkmarks, arrows, etc.) convert to
# empty boxes in the PDF instead of just looking visually different.
COPY fontconfig-aliases.conf /etc/fonts/local.conf
RUN fc-cache -f

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
