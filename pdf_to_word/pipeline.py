"""
PDF → DOCX conversion pipeline orchestrator.

Coordinates all stages in the correct order and surfaces clean, user-safe
errors (InvalidFileError) rather than raw library exceptions.

Conversion flow
---------------
  1. Validate + triage (password? corrupted? scanned?)
  2. Per-page: extract geometry + raw blocks → layout_extractor
  3. Cross-page: detect running headers/footers → element_classifier
  4. Per-page: sort reading order + classify headings + lists
  5. Assemble DocumentModel → docx_builder → save

Bug 4 fix: fitz.open() is always used as a context manager (`with` block)
so the OS file descriptor is released deterministically — even on Windows,
where an open handle would lock the file and cause sharing-violation errors
in batch conversions or when the caller tries to move/delete the file after.
"""
from __future__ import annotations

import os

import fitz  # PyMuPDF

from .analyzer.element_classifier import ElementClassifier
from .analyzer.layout_extractor import LayoutExtractor
from .analyzer.pdf_scanner import PDFScanner
from .builder.docx_builder import DocxBuilder
from .handlers.image_handler import ImageHandler
from .handlers.table_handler import TableHandler
from .models.document_schema import DocumentModel, DocumentPage, PageGeometry, Block


class InvalidFileError(ValueError):
    """Raised for user-correctable file problems (bad PDF, password, etc.)."""


class ConversionResult:
    """Returned by convert() so callers know which path was taken."""
    def __init__(self, output_path: str, method: str, is_scanned: bool) -> None:
        self.output_path = output_path
        self.method      = method       # "digital" | "scanned-image-fallback"
        self.is_scanned  = is_scanned


def convert(pdf_path: str, output_path: str) -> ConversionResult:
    """
    Main entry point.  Raises InvalidFileError for user-correctable problems.
    All other exceptions propagate as-is for the caller to wrap in HTTP 500.
    """
    # --- 0. Pre-flight validation (file must exist, be a valid PDF) ---
    if not os.path.isfile(pdf_path):
        raise InvalidFileError(f"File not found: {pdf_path}")

    ok, reason = PDFScanner.is_valid_pdf(pdf_path)
    if not ok:
        _raise_for_reason(reason)

    # --- 1. Open once, do everything inside the context manager ---
    # (Bug 4 fix: with-block guarantees close() on exit and exception)
    with fitz.open(pdf_path) as doc:

        # 1a. Scanned check
        is_scanned, _scanned_pages = PDFScanner.analyse(doc)
        if is_scanned:
            # Scanned fallback: embed each page as a full-page picture.
            # OCR engine will plug in here in a later step.
            _build_scanned_fallback(doc, output_path)
            return ConversionResult(
                output_path = output_path,
                method      = "scanned-image-fallback",
                is_scanned  = True,
            )

        extractor = LayoutExtractor(doc)
        table_handler = TableHandler()
        image_handler = ImageHandler(doc)
        doc_model = DocumentModel()

        # Collect (blocks, geometry) per page for cross-page analysis
        page_data: list[tuple[list[Block], PageGeometry]] = []

        # 2. Per-page extraction
        for page in doc:
            doc_page: DocumentPage = extractor.extract_page(page)
            drawings = extractor.get_drawings(page.number)

            # 2a. Detect tables from grid lines and suppress table paragraphs
            blocks_after_tables, remaining_drawings = table_handler.process_page_tables(
                page=page,
                existing_blocks=doc_page.blocks,
                raw_drawings=drawings,
            )

            # 2b. Detect vector diagrams & raster images using non-table drawings
            doc_page.blocks = image_handler.process_page_visuals(
                page=page,
                existing_blocks=blocks_after_tables,
                raw_drawings=remaining_drawings,
            )
            doc_model.pages.append(doc_page)
            page_data.append((doc_page.blocks, doc_page.geometry))

        # 3. Cross-page: mark running headers / footers
        ElementClassifier.detect_headers_footers(page_data)

        # 4. Per-page: reading order + semantic classification
        for doc_page in doc_model.pages:
            doc_page.blocks = ElementClassifier.sort_reading_order(doc_page.blocks)
            ElementClassifier.classify_headings(doc_page.blocks)
            ElementClassifier.classify_lists(doc_page.blocks)

        # Metadata from PDF info dict
        info = doc.metadata or {}
        doc_model.title  = info.get("title")  or None
        doc_model.author = info.get("author") or None

    # fitz.Document is now closed — all fitz.Page references are invalid
    # past this point, which is fine since we only use doc_model from here.

    # 5. Build DOCX
    builder = DocxBuilder(doc_model)
    builder.build(output_path)

    return ConversionResult(
        output_path = output_path,
        method      = "digital",
        is_scanned  = False,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _raise_for_reason(reason: str) -> None:
    messages = {
        "password_protected": (
            "This PDF is password-protected. Please remove the password "
            "(e.g. via your PDF viewer's 'Print to PDF' option) and try again."
        ),
        "corrupted_or_not_pdf": (
            "This doesn't look like a valid PDF file. It may be corrupted or "
            "not actually a PDF — please check the file and try again."
        ),
        "empty_document": (
            "The PDF appears to contain no pages."
        ),
    }
    raise InvalidFileError(messages.get(reason, f"Could not open PDF: {reason}"))


def _build_scanned_fallback(doc: fitz.Document, output_path: str) -> None:
    """
    Embeds each page as a 200-DPI PNG inside a Word document.
    Not editable text, but at least the visual content is preserved.
    OCR engine will replace this path in a later pipeline step.
    """
    from io import BytesIO
    from docx import Document
    from docx.shared import Inches

    out = Document()
    section = out.sections[0]
    usable_w = (
        section.page_width.inches
        - section.left_margin.inches
        - section.right_margin.inches
    )

    for i, page in enumerate(doc):
        if i > 0:
            out.add_page_break()
        pix = page.get_pixmap(dpi=200)
        out.add_picture(BytesIO(pix.tobytes("png")), width=Inches(usable_w))

    out.save(output_path)
