"""
PDF type triage: determines whether a document needs the OCR path
(scanned image PDF) or the direct text-extraction path (digital PDF).

Design decisions:
  • Per-page character count, not total — avoids misclassifying a
    hybrid doc (e.g. 8 digital pages + 1 scanned page) as fully digital.
  • Threshold is intentionally low (20 chars) because even scanned PDFs
    often contain a handful of extractable characters from embedded
    metadata, page numbers rendered as actual text, or a partial text
    layer left by the scanner software.
  • Returns (is_scanned: bool, scanned_pages: list[int]) so callers can
    decide whether to run OCR only on the bad pages or reject the whole
    document.
"""
from __future__ import annotations

import fitz  # PyMuPDF


class PDFScanner:
    # A page is considered "image-only / no text" if it has fewer than
    # this many extractable characters after stripping whitespace.
    _CHARS_PER_PAGE_THRESHOLD: int = 20

    # If at least this fraction of pages are text-poor, treat the whole
    # document as scanned (rather than requiring ALL pages to fail).
    _SCANNED_PAGE_RATIO: float = 0.85

    @classmethod
    def analyse(
        cls,
        doc: fitz.Document,
    ) -> tuple[bool, list[int]]:
        """
        Returns:
            (is_scanned, scanned_page_indices)

        is_scanned is True when >= _SCANNED_PAGE_RATIO fraction of pages
        have fewer than _CHARS_PER_PAGE_THRESHOLD extractable characters.
        scanned_page_indices is the list of 0-based page numbers that
        individually failed the threshold (useful for hybrid-doc handling).
        """
        total = len(doc)
        if total == 0:
            return False, []

        poor_pages: list[int] = []
        for page in doc:
            char_count = len(page.get_text("text").strip())
            if char_count < cls._CHARS_PER_PAGE_THRESHOLD:
                poor_pages.append(page.number)

        ratio = len(poor_pages) / total
        is_scanned = ratio >= cls._SCANNED_PAGE_RATIO
        return is_scanned, poor_pages

    @classmethod
    def is_password_protected(cls, doc: fitz.Document) -> bool:
        """
        fitz.open() on an encrypted PDF succeeds but needs_pass == True.
        Checking this before extraction avoids cryptic downstream errors.
        """
        return doc.needs_pass

    @classmethod
    def is_valid_pdf(cls, path: str) -> tuple[bool, str]:
        """
        Opens and immediately closes the file to verify it is a valid,
        non-encrypted PDF.  Returns (ok, reason_string).
        """
        try:
            with fitz.open(path) as doc:
                if cls.is_password_protected(doc):
                    return False, "password_protected"
                if len(doc) == 0:
                    return False, "empty_document"
            return True, ""
        except fitz.FileDataError:
            return False, "corrupted_or_not_pdf"
        except Exception as exc:
            return False, str(exc)
