"""
Visual content handler: raster images and vector diagram detection.

Responsibilities:
  1. Detect dense clusters of vector drawing paths (flowcharts, diagrams,
     charts rendered as PDF line-art) and rasterize them at 300 DPI.
  2. Suppress text blocks that fall INSIDE a detected vector diagram's
     bounding box — those labels are already baked into the rasterized
     image, so emitting them as separate Word paragraphs would duplicate
     them (once in the picture, once as selectable text on top of it).

Raster images (embedded PNGs/JPEGs) are already extracted by
layout_extractor.py's _parse_image_block — this handler only adds the
vector-diagram path that layout_extractor intentionally doesn't handle
(vector paths aren't "blocks" in PyMuPDF's text-dict output).

Design decisions:
  • Iterative merge for vector clusters: a single-pass union misses
    transitive overlaps (A touches C, C touches B, but A and B don't
    touch each other directly).  The loop repeats until no more merges
    happen.
  • Minimum cluster size (40×40 pt) filters out underlines, borders,
    and thin separator rules that are technically "drawings" but not
    diagrams.
  • 300 DPI rendering (zoom = 300/72 ≈ 4.17) balances quality vs. file
    size — 200 DPI is too blurry for small text inside diagrams, 600 DPI
    inflates the .docx without visible benefit.
"""
from __future__ import annotations

import fitz  # PyMuPDF

from ..models.document_schema import Block, BoundingBox, ElementType

# Diagrams smaller than this (in points) are probably decorative lines.
_MIN_CLUSTER_WIDTH:  float = 40.0
_MIN_CLUSTER_HEIGHT: float = 40.0

# Padding (pt) used to test whether two drawing rects are "adjacent enough"
# to belong to the same diagram cluster.
_CLUSTER_PAD: float = 6.0

# Render resolution for vector diagram rasterization.
_RENDER_DPI: int = 300


class ImageHandler:
    def __init__(self, doc: fitz.Document) -> None:
        self.doc = doc

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def process_page_visuals(
        self,
        page: fitz.Page,
        existing_blocks: list[Block],
        raw_drawings: list[dict],
    ) -> list[Block]:
        """
        Detects vector diagram clusters on *page*, rasterizes them at
        300 DPI, removes text blocks that fall inside those clusters
        (already baked into the image), and returns the merged result.

        Parameters
        ----------
        page : fitz.Page
            The PDF page to render diagram crops from.
        existing_blocks : list[Block]
            Blocks already extracted by LayoutExtractor (text + raster images).
        raw_drawings : list[dict]
            Drawing primitives from page.get_drawings(), cached by
            LayoutExtractor.drawings_cache[page_num].

        Returns
        -------
        list[Block]
            *existing_blocks* minus suppressed text blocks, plus new
            VECTOR_DIAGRAM blocks.
        """
        # 1. Detect vector diagram clusters
        clusters = self._find_vector_clusters(raw_drawings, page.number)

        if not clusters:
            return existing_blocks

        # 2. Find text blocks inside each cluster → suppress them
        suppressed: set[int] = set()
        diagram_blocks: list[Block] = []

        for cluster_box in clusters:
            # Skip tiny clusters (borders, underlines)
            if (cluster_box.width < _MIN_CLUSTER_WIDTH
                    or cluster_box.height < _MIN_CLUSTER_HEIGHT):
                continue

            # Mark contained text blocks for removal
            for idx, blk in enumerate(existing_blocks):
                if (blk.element_type == ElementType.PARAGRAPH
                        and self._is_inside(blk.bbox, cluster_box)):
                    suppressed.add(idx)

            # 3. Rasterize the cluster region at 300 DPI
            img_bytes = self._rasterize_region(page, cluster_box)
            if img_bytes:
                diagram_blocks.append(
                    Block(
                        element_type = ElementType.VECTOR_DIAGRAM,
                        bbox         = cluster_box,
                        image_bytes  = img_bytes,
                        image_format = "png",
                    )
                )

        # 4. Remove suppressed text blocks and append diagram blocks
        clean_blocks = [
            blk for idx, blk in enumerate(existing_blocks)
            if idx not in suppressed
        ]
        return clean_blocks + diagram_blocks

    # ------------------------------------------------------------------
    # Vector cluster detection
    # ------------------------------------------------------------------

    def _find_vector_clusters(
        self,
        drawings: list[dict],
        page_num: int,
    ) -> list[BoundingBox]:
        """
        Groups nearby drawing primitives into diagram clusters via
        iterative bounding-box merging.

        Bug 2 fix: the merge loop repeats until no more unions happen,
        so transitive overlaps (A→C→B) are caught even if A and B don't
        directly touch.
        """
        if not drawings:
            return []

        # Extract raw rects from drawing dicts
        rects: list[fitz.Rect] = []
        for d in drawings:
            raw_rect = d.get("rect")
            if raw_rect:
                r = fitz.Rect(raw_rect)
                if not r.is_empty and not r.is_infinite:
                    rects.append(r)

        if not rects:
            return []

        # --- Iterative merge until stable ---
        merged = list(rects)
        changed = True

        while changed:
            changed = False
            new_merged: list[fitz.Rect] = []

            while merged:
                current = merged.pop(0)
                found_overlap = False

                for i, other in enumerate(new_merged):
                    # Padded intersection check
                    expanded = fitz.Rect(
                        other.x0 - _CLUSTER_PAD,
                        other.y0 - _CLUSTER_PAD,
                        other.x1 + _CLUSTER_PAD,
                        other.y1 + _CLUSTER_PAD,
                    )
                    if current.intersects(expanded):
                        new_merged[i] = other | current  # union
                        found_overlap = True
                        changed = True
                        break

                if not found_overlap:
                    new_merged.append(current)

            merged = new_merged

        return [
            BoundingBox(
                x0=r.x0, y0=r.y0, x1=r.x1, y1=r.y1,
                page_num=page_num,
            )
            for r in merged
        ]

    # ------------------------------------------------------------------
    # Rasterization
    # ------------------------------------------------------------------

    @staticmethod
    def _rasterize_region(
        page: fitz.Page,
        box: BoundingBox,
    ) -> bytes | None:
        """Render the given bounding box at _RENDER_DPI as a PNG."""
        clip = fitz.Rect(box.x0, box.y0, box.x1, box.y1)
        if clip.is_empty:
            return None
        try:
            zoom = _RENDER_DPI / 72.0
            mat  = fitz.Matrix(zoom, zoom)
            pix  = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
            return pix.tobytes("png")
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Geometry helper
    # ------------------------------------------------------------------

    @staticmethod
    def _is_inside(inner: BoundingBox, outer: BoundingBox) -> bool:
        """True when *inner* falls entirely within *outer* (±2 pt tolerance)."""
        return (
            inner.x0 >= (outer.x0 - 2)
            and inner.y0 >= (outer.y0 - 2)
            and inner.x1 <= (outer.x1 + 2)
            and inner.y1 <= (outer.y1 + 2)
        )
