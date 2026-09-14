"""
Raw geometry and primitive extraction from a PDF page via PyMuPDF.

Responsibilities:
  1. Page dimensions + margin estimation  → PageGeometry
  2. Text blocks with per-span metadata   → list[Block] (PARAGRAPH / IMAGE)
  3. Vector drawing primitives cache      → stored for table_handler to reuse

Key design decisions:
  • Gap-based space insertion (not blind text+" "): a space run is inserted
    between consecutive spans only when there is a measurable horizontal gap
    AND the adjacent characters are not punctuation that naturally attaches
    without a space (e.g. commas, closing brackets).
  • TEXT_DEHYPHENATE flag: PyMuPDF re-joins soft-hyphenated words that were
    split across lines at the PDF level, so "infor-\nmation" becomes
    "information" before we ever see it.
  • Drawings are extracted once per page and stored in self.drawings_cache
    so table_handler (Step 6) can read them without re-opening the file.
"""
from __future__ import annotations

import fitz  # PyMuPDF

from ..models.document_schema import (
    Block,
    BoundingBox,
    DocumentPage,
    ElementType,
    PageGeometry,
    TextSpan,
)

# Characters that should NOT have a space inserted immediately before them.
_NO_SPACE_BEFORE: frozenset[str] = frozenset(".,;:!?)]}%\"'\u2019\u201d\u2014\u2013")
# Characters that should NOT have a space inserted immediately after them.
_NO_SPACE_AFTER: frozenset[str] = frozenset("([{\"'\u2018\u201c$#@\u2014\u2013")

# Minimum horizontal gap (in points) between two spans to insert a space.
_SPACE_GAP_THRESHOLD: float = 1.5


class LayoutExtractor:
    def __init__(self, doc: fitz.Document) -> None:
        self.doc = doc
        # page_num (0-based) → list of raw drawing dicts from get_drawings()
        # Populated lazily on first access per page.
        self.drawings_cache: dict[int, list[dict]] = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract_page(self, page: fitz.Page) -> DocumentPage:
        """
        Full extraction for one page: geometry + text blocks + image blocks.
        Drawings are cached but NOT returned here — table_handler reads them
        via self.drawings_cache[page_num].
        """
        geometry = self._extract_geometry(page)
        self._cache_drawings(page)
        blocks = self._extract_blocks(page, geometry)
        return DocumentPage(geometry=geometry, blocks=blocks)

    def get_drawings(self, page_num: int) -> list[dict]:
        """Return cached drawing primitives for the given page (0-based)."""
        return self.drawings_cache.get(page_num, [])

    # ------------------------------------------------------------------
    # Geometry
    # ------------------------------------------------------------------

    def _extract_geometry(self, page: fitz.Page) -> PageGeometry:
        rect = page.rect

        # Margin estimation: find the bounding box of all text content on the
        # page and use the distance from the page edge as the margin proxy.
        # Falls back to a safe 0.5-inch (36 pt) default when no text exists.
        content_rect = self._estimate_content_rect(page, rect)

        margin_left   = max(18.0, content_rect.x0)
        margin_top    = max(18.0, content_rect.y0)
        margin_right  = max(18.0, rect.width  - content_rect.x1)
        margin_bottom = max(18.0, rect.height - content_rect.y1)

        return PageGeometry(
            page_num      = page.number,
            width         = rect.width,
            height        = rect.height,
            margin_top    = round(margin_top,    1),
            margin_bottom = round(margin_bottom, 1),
            margin_left   = round(margin_left,   1),
            margin_right  = round(margin_right,  1),
        )

    def _estimate_content_rect(
        self, page: fitz.Page, page_rect: fitz.Rect
    ) -> fitz.Rect:
        """
        Union of all text-line bounding boxes.  Gives a real content boundary
        rather than hard-coding a margin value.
        """
        x0s, y0s, x1s, y1s = [], [], [], []
        data = page.get_text("dict", flags=fitz.TEXT_DEHYPHENATE)
        for blk in data.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                b = line.get("bbox", [])
                if len(b) == 4:
                    x0s.append(b[0]); y0s.append(b[1])
                    x1s.append(b[2]); y1s.append(b[3])

        if not x0s:
            # Fallback: 0.5-inch margins
            return fitz.Rect(36, 36, page_rect.width - 36, page_rect.height - 36)

        return fitz.Rect(min(x0s), min(y0s), max(x1s), max(y1s))

    # ------------------------------------------------------------------
    # Drawings cache
    # ------------------------------------------------------------------

    def _cache_drawings(self, page: fitz.Page) -> None:
        if page.number not in self.drawings_cache:
            try:
                self.drawings_cache[page.number] = page.get_drawings()
            except Exception:
                self.drawings_cache[page.number] = []

    # ------------------------------------------------------------------
    # Block extraction
    # ------------------------------------------------------------------

    def _extract_blocks(
        self, page: fitz.Page, geometry: PageGeometry
    ) -> list[Block]:
        blocks: list[Block] = []
        page_dict = page.get_text("dict", flags=fitz.TEXT_DEHYPHENATE)

        for raw_b in page_dict.get("blocks", []):
            bbox = self._make_bbox(raw_b["bbox"], page.number)

            if raw_b.get("type") == 0:          # text block
                block = self._parse_text_block(raw_b, bbox)
                if block is not None:
                    blocks.append(block)

            elif raw_b.get("type") == 1:         # inline image
                block = self._parse_image_block(raw_b, page, bbox)
                if block is not None:
                    blocks.append(block)

        return blocks

    # ------------------------------------------------------------------
    # Text block parsing
    # ------------------------------------------------------------------

    def _parse_text_block(self, raw_b: dict, bbox: BoundingBox) -> Block | None:
        spans: list[TextSpan] = []

        for line in raw_b.get("lines", []):
            line_spans = self._parse_line_spans(line)
            spans.extend(line_spans)

        if not spans:
            return None

        return Block(
            element_type = ElementType.PARAGRAPH,
            bbox         = bbox,
            spans        = spans,
        )

    def _parse_line_spans(self, line: dict) -> list[TextSpan]:
        """
        Converts one PDF text line into TextSpan objects.

        Space insertion rule (Bug 1 fix):
          A space span is added between consecutive raw spans only when:
            1. There is a measurable horizontal gap (> _SPACE_GAP_THRESHOLD pt)
            2. The last character of the previous text is not in _NO_SPACE_AFTER
            3. The first character of the next text is not in _NO_SPACE_BEFORE
        This prevents phantom spaces inside compound tokens (e.g. "don't",
        "$50.00") while still separating genuinely distinct words.
        """
        result: list[TextSpan] = []
        raw_spans = line.get("spans", [])
        last_x1: float | None = None
        last_text: str = ""

        for s in raw_spans:
            raw_text = s.get("text", "")
            if not raw_text:
                continue

            sbbox = s.get("bbox", [0.0, 0.0, 0.0, 0.0])
            x0 = sbbox[0]

            # --- Gap-based space insertion ---
            if last_x1 is not None:
                gap = x0 - last_x1
                if gap > _SPACE_GAP_THRESHOLD:
                    prev_char = last_text[-1] if last_text else ""
                    next_char = raw_text[0]  if raw_text  else ""
                    if (
                        prev_char not in _NO_SPACE_AFTER
                        and next_char not in _NO_SPACE_BEFORE
                    ):
                        # Inherit font attrs from the current span for the space
                        result.append(
                            TextSpan(
                                text      = " ",
                                font_name = s.get("font", "Arial"),
                                font_size = round(s.get("size", 11.0), 1),
                            )
                        )

            flags    = s.get("flags", 0)
            is_bold  = bool(flags & 16)   # bit 4
            is_italic = bool(flags & 2)   # bit 1

            color_int = s.get("color", 0)
            # Ensure 6-digit hex, no '#' prefix (matches TextSpan contract)
            color_hex = f"{color_int & 0xFFFFFF:06x}"

            result.append(
                TextSpan(
                    text      = raw_text,
                    font_name = s.get("font", "Arial"),
                    font_size = round(s.get("size", 11.0), 1),
                    is_bold   = is_bold,
                    is_italic = is_italic,
                    color_hex = color_hex,
                )
            )

            last_x1   = sbbox[2]
            last_text = raw_text

        return result

    # ------------------------------------------------------------------
    # Image block parsing
    # ------------------------------------------------------------------

    def _parse_image_block(
        self, raw_b: dict, page: fitz.Page, bbox: BoundingBox
    ) -> Block | None:
        """
        Tries to extract a high-quality PNG of this image block.
        Falls back to the inline bytes already in the dict if pixmap
        extraction fails (e.g. the image is part of a form XObject).
        """
        img_bytes: bytes | None = None
        img_format: str = "png"

        # Prefer extracting via xref for proper colour space handling
        # (inline dict images may be CMYK or have an alpha mask that
        # fitz doesn't resolve automatically).
        xref: int = raw_b.get("number", 0)
        if xref > 0:
            try:
                pix = fitz.Pixmap(self.doc, xref)
                if pix.n - pix.alpha >= 4:   # CMYK → RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                img_bytes  = pix.tobytes("png")
                img_format = "png"
            except Exception:
                img_bytes = None

        if img_bytes is None:
            # Fallback: use whatever bytes the dict already has
            img_bytes = raw_b.get("image", None)
            img_format = raw_b.get("ext", "png")

        if not img_bytes:
            return None

        # Skip tiny decorative blobs (< 5×5 pt)
        if bbox.width < 5 or bbox.height < 5:
            return None

        return Block(
            element_type = ElementType.IMAGE,
            bbox         = bbox,
            image_bytes  = img_bytes,
            image_format = img_format,
        )

    # ------------------------------------------------------------------
    # Helper
    # ------------------------------------------------------------------

    @staticmethod
    def _make_bbox(raw: list | tuple, page_num: int) -> BoundingBox:
        return BoundingBox(
            x0=raw[0], y0=raw[1], x1=raw[2], y1=raw[3],
            page_num=page_num,
        )
