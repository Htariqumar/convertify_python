FROM python:3.11-slim

# LibreOffice (word/excel/ppt <-> pdf, office thumbnails) and Ghostscript (compress-pdf,
# pdf thumbnails) - both installed as normal Debian packages since this runs on a regular
# Linux host, not a size-capped serverless function. fonts-* packages keep converted
# documents from silently substituting missing glyphs.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    ghostscript \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8001
EXPOSE 8001

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
