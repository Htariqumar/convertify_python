"""Standalone entry point for the pdf2docx conversion step, run as a subprocess by
main.py's pdf-to-word endpoint (see _run_pdf2docx_conversion) instead of calling
pdf2docx in-process. pdf2docx is a third-party library with known cases where it
hangs indefinitely on certain malformed/pathological PDFs; running it here lets the
caller enforce a hard wall-clock timeout via subprocess.run(timeout=...), which kills
this whole process (and pdf2docx's C-level work with it) if it's exceeded. An
in-process thread can't be forcibly killed the same way, so a hang there would leak
a stuck thread-pool worker for good - and since the service runs as a single process
(see UVICORN_WORKERS in the Dockerfile), enough of those eventually exhaust the pool
or its memory and take every conversion tool down for every user.

Deliberately not just `python -c "..."` inline in main.py: this must NOT import
main.py itself, which would re-run its heavy module-level setup (loading the Whisper
model, etc.) on every single pdf-to-word request.
"""
import sys

from pdf2docx import Converter


def main() -> int:
    pdf_path, docx_path = sys.argv[1], sys.argv[2]
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path, start=0, end=None)
    finally:
        cv.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
