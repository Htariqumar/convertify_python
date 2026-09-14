"""
Block classification and reading-order correction.

Two independent responsibilities, deliberately kept as static methods so
callers (pipeline.py) can compose them freely:

  1. sort_reading_order()   — re-orders blocks into human reading sequence
     (Bug 2 fix: tracks column x-interval rather than a single x0 point,
      so indented paragraphs stay in their correct column)

  2. classify_headings()    — upgrades PARAGRAPH blocks to HEADING based on
     font size relative to the document's *actual* median body font size
     (not a hard-coded 11 pt constant that breaks non-standard documents)

  3. classify_lists()       — detects bullet/numbered list items from the
     first non-space character of the block's plain text

  4. detect_header_footer() — marks blocks that fall within the top/bottom
     margin bands and repeat across pages as HEADER / FOOTER so the builder
     can place them in Word's real header/footer area instead of the body
"""
from __future__ import annotations

import re
import statistics
from collections import Counter

from ..models.document_schema import Block, ElementType, PageGeometry


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# A column "owns" a new block if the block's x0 falls within this many points
# of the column's recorded x-interval [x_min, x_max].
_COL_GAP_THRESHOLD: float = 30.0

# Fraction of the page height that defines the header / footer zone.
_HEADER_ZONE_RATIO: float = 0.10
_FOOTER_ZONE_RATIO: float = 0.10

# A text pattern must repeat on at least this fraction of pages to be
# classified as a running header/footer.
_REPEAT_THRESHOLD_RATIO: float = 0.60
_REPEAT_MIN_PAGES: int = 2       # never classify on a single-page doc


class ElementClassifier:

    # ------------------------------------------------------------------
    # 1. Reading order
    # ------------------------------------------------------------------

    @staticmethod
    def sort_reading_order(
        blocks: list[Block],
        col_gap_threshold: float = _COL_GAP_THRESHOLD,
    ) -> list[Block]:
        """
        Re-orders blocks into correct human reading sequence for both
        single-column and multi-column layouts.

        Algorithm:
          a) Pre-sort by (y0, x0) so blocks arrive roughly top-to-bottom.
          b) Assign each block to a "column bucket" by checking whether
             its x0 falls within an existing column's [x_min, x_max]
             interval (±col_gap_threshold).  If not, a new column is created.
             Tracking the interval (not a single x0) means indented
             paragraphs stay in the same column as their heading.
          c) Sort column buckets left-to-right by their x_min.
          d) Within each column, sort top-to-bottom by y0.
          e) Assign sequential order_index values.

        Limitation: assumes at most 3 columns (sufficient for academic /
        magazine layouts; a real newspaper with 6 columns would need a
        more sophisticated gap-analysis approach).
        """
        if not blocks:
            return []

        # (a) coarse top-to-bottom pre-sort
        pre_sorted = sorted(blocks, key=lambda b: (b.bbox.y0, b.bbox.x0))

        # (b) assign to column buckets
        # Each bucket: {"blocks": [...], "x_min": float, "x_max": float}
        columns: list[dict] = []

        for blk in pre_sorted:
            bx0, bx1 = blk.bbox.x0, blk.bbox.x1
            placed = False

            for col in columns:
                # Check if block x0 falls within this column's x-interval
                col_lo = col["x_min"] - col_gap_threshold
                col_hi = col["x_max"] + col_gap_threshold
                if col_lo <= bx0 <= col_hi:
                    col["blocks"].append(blk)
                    col["x_min"] = min(col["x_min"], bx0)
                    col["x_max"] = max(col["x_max"], bx1)
                    placed = True
                    break

            if not placed:
                columns.append({"blocks": [blk], "x_min": bx0, "x_max": bx1})

        # (c) sort columns left-to-right
        columns.sort(key=lambda c: c["x_min"])

        # (d+e) flatten + assign order_index
        ordered: list[Block] = []
        for col in columns:
            col_blocks = sorted(col["blocks"], key=lambda b: b.bbox.y0)
            ordered.extend(col_blocks)

        for idx, blk in enumerate(ordered):
            blk.order_index = idx

        return ordered

    # ------------------------------------------------------------------
    # 2. Heading classification
    # ------------------------------------------------------------------

    @staticmethod
    def classify_headings(blocks: list[Block]) -> None:
        """
        Promotes PARAGRAPH blocks to HEADING based on font size relative
        to the document's median body font size (computed dynamically).

        Heading level rules (relative to median body_size):
          H1 : font_size >= body_size + 5
          H2 : font_size >= body_size + 2.5
          H3 : bold AND font_size >= body_size  (styled bold at body size)

        Modifies blocks in-place; returns None.
        """
        # Collect all font sizes from paragraph spans to find true body size
        all_sizes: list[float] = [
            s.font_size
            for b in blocks
            if b.element_type == ElementType.PARAGRAPH
            for s in b.spans
            if s.text.strip()
        ]
        if not all_sizes:
            return

        body_size = statistics.median(all_sizes)

        for blk in blocks:
            if blk.element_type != ElementType.PARAGRAPH or not blk.spans:
                continue

            max_size = max(s.font_size for s in blk.spans)
            is_bold  = any(s.is_bold for s in blk.spans)
            text     = blk.plain_text.strip()

            # Skip very long blocks — they are body paragraphs, not headings
            if len(text) > 200:
                continue

            if max_size >= body_size + 5:
                blk.element_type  = ElementType.HEADING
                blk.heading_level = 1
            elif max_size >= body_size + 2.5:
                blk.element_type  = ElementType.HEADING
                blk.heading_level = 2
            elif is_bold and max_size >= body_size and len(text) < 80:
                blk.element_type  = ElementType.HEADING
                blk.heading_level = 3

    # ------------------------------------------------------------------
    # 3. List detection
    # ------------------------------------------------------------------

    # Matches: •, ◦, ▪, ▸, –, -, * followed by whitespace
    _BULLET_RE = re.compile(r"^[\u2022\u25e6\u25aa\u25b8\u2013\-\*]\s")
    # Matches: "1.", "2)", "a.", "(i)" etc. followed by whitespace
    _NUMBERED_RE = re.compile(r"^(\(?\w{1,3}[\.\)]\s)")

    @classmethod
    def classify_lists(cls, blocks: list[Block]) -> None:
        """
        Promotes PARAGRAPH blocks to LIST_ITEM when the block's plain text
        starts with a recognised bullet or numbered-list marker.
        Modifies blocks in-place.
        """
        for blk in blocks:
            if blk.element_type != ElementType.PARAGRAPH:
                continue
            text = blk.plain_text.lstrip()
            if cls._BULLET_RE.match(text):
                blk.element_type = ElementType.LIST_ITEM
                blk.list_type    = "bullet"
            elif cls._NUMBERED_RE.match(text):
                blk.element_type = ElementType.LIST_ITEM
                blk.list_type    = "numbered"

    # ------------------------------------------------------------------
    # 4. Header / footer detection  (cross-page, must be called once
    #    after all pages are extracted)
    # ------------------------------------------------------------------

    @staticmethod
    def detect_headers_footers(
        all_page_blocks: list[tuple[list[Block], PageGeometry]],
    ) -> None:
        """
        Marks PARAGRAPH blocks that appear repeatedly in the top or bottom
        margin band of pages as HEADER or FOOTER.

        Accepts a list of (blocks_for_page, geometry_for_page) tuples in
        page order.  Modifies blocks in-place.

        Pattern matching: digits in the text are normalised to '#' before
        comparison so "Page 1", "Page 2" … match as the same pattern.
        """
        total_pages = len(all_page_blocks)
        if total_pages < _REPEAT_MIN_PAGES:
            return

        repeat_threshold = max(
            _REPEAT_MIN_PAGES,
            int(total_pages * _REPEAT_THRESHOLD_RATIO),
        )

        top_counter:    Counter[str] = Counter()
        bottom_counter: Counter[str] = Counter()
        # pattern → list of (page_idx, block) for targeted marking
        top_hits:    dict[str, list[tuple[int, Block]]] = {}
        bottom_hits: dict[str, list[tuple[int, Block]]] = {}

        # --- Pass 1: count patterns ---
        for page_idx, (blocks, geo) in enumerate(all_page_blocks):
            header_y1 = geo.height * _HEADER_ZONE_RATIO
            footer_y0 = geo.height * (1.0 - _FOOTER_ZONE_RATIO)

            for blk in blocks:
                if blk.element_type != ElementType.PARAGRAPH:
                    continue
                text = blk.plain_text.strip()
                if not text:
                    continue
                pattern = re.sub(r"\d+", "#", text)

                if blk.bbox.y1 <= header_y1:
                    top_counter[pattern] += 1
                    top_hits.setdefault(pattern, []).append((page_idx, blk))
                elif blk.bbox.y0 >= footer_y0:
                    bottom_counter[pattern] += 1
                    bottom_hits.setdefault(pattern, []).append((page_idx, blk))

        # --- Pass 2: mark qualifying patterns ---
        for pattern, count in top_counter.items():
            if count >= repeat_threshold:
                for _, blk in top_hits[pattern]:
                    blk.element_type = ElementType.HEADER

        for pattern, count in bottom_counter.items():
            if count >= repeat_threshold:
                for _, blk in bottom_hits[pattern]:
                    blk.element_type = ElementType.FOOTER
