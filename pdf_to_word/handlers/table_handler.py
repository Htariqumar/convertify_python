"""
Table extraction handler for PDF → Word conversion.

Responsibilities:
  1. Extract horizontal and vertical vector lines (from path items and thin rects).
  2. Cluster intersecting lines into distinct table components (supports multiple tables per page).
  3. Reconstruct grid cells (rows and columns matrix).
  4. Assign text blocks/spans to cells using robust center-point containment.
  5. Suppress contained text blocks from the document body flow so text isn't duplicated.
  6. Expose table bounding boxes so ImageHandler does not mistake table borders for vector diagrams.
"""
from __future__ import annotations

import fitz  # PyMuPDF

from ..models.document_schema import (
    Block,
    BoundingBox,
    ElementType,
    TableCellData,
    TextSpan,
)

# Minimum dimensions (in points) to consider a grid a valid table
_MIN_TABLE_WIDTH: float = 50.0
_MIN_TABLE_HEIGHT: float = 25.0
_MIN_LINE_LENGTH: float = 15.0


class TableHandler:
    def __init__(self, line_tolerance: float = 3.0) -> None:
        self.line_tolerance = line_tolerance

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def process_page_tables(
        self,
        page: fitz.Page,
        existing_blocks: list[Block],
        raw_drawings: list[dict],
    ) -> tuple[list[Block], list[dict]]:
        """
        Detects tables from raw_drawings, builds Block(TABLE) with table_matrix,
        suppresses contained text blocks, and filters out drawings that belong
        to tables so ImageHandler won't turn table lines into images.

        Returns
        -------
        (blocks, remaining_drawings)
            blocks: body blocks with table blocks inserted and contained paragraphs suppressed.
            remaining_drawings: drawings minus lines that belong to tables.
        """
        h_segments, v_segments = self._extract_line_segments(raw_drawings)

        if len(h_segments) < 2 or len(v_segments) < 2:
            return existing_blocks, raw_drawings

        table_grids = self._cluster_grids(h_segments, v_segments, page.number)
        if not table_grids:
            return existing_blocks, raw_drawings

        table_blocks: list[Block] = []
        suppressed_block_indices: set[int] = set()
        table_bboxes: list[BoundingBox] = []

        for grid in table_grids:
            table_bbox = grid["bbox"]
            table_bboxes.append(table_bbox)
            rows_coords: list[float] = grid["rows"]
            cols_coords: list[float] = grid["cols"]

            matrix: list[list[TableCellData]] = []

            for r_idx in range(len(rows_coords) - 1):
                row_cells: list[TableCellData] = []
                y0, y1 = rows_coords[r_idx], rows_coords[r_idx + 1]

                for c_idx in range(len(cols_coords) - 1):
                    x0, x1 = cols_coords[c_idx], cols_coords[c_idx + 1]
                    cell_bbox = BoundingBox(
                        x0=x0, y0=y0, x1=x1, y1=y1, page_num=page.number
                    )

                    # Collect blocks whose center point falls inside this cell
                    matched_blocks: list[Block] = []
                    for b_idx, b in enumerate(existing_blocks):
                        if b.element_type == ElementType.PARAGRAPH:
                            if self._is_cell_content(b.bbox, cell_bbox):
                                matched_blocks.append(b)
                                suppressed_block_indices.add(b_idx)

                    # Sort matched blocks top-to-bottom, left-to-right
                    matched_blocks.sort(key=lambda blk: (blk.bbox.y0, blk.bbox.x0))

                    cell_spans: list[TextSpan] = []
                    for mb in matched_blocks:
                        if cell_spans and mb.spans:
                            # Add a space between separate paragraphs inside cell
                            cell_spans.append(TextSpan(text=" "))
                        cell_spans.extend(mb.spans)

                    row_cells.append(
                        TableCellData(
                            bbox=cell_bbox,
                            spans=cell_spans,
                            row_span=1,
                            col_span=1,
                        )
                    )
                matrix.append(row_cells)

            table_blocks.append(
                Block(
                    element_type=ElementType.TABLE,
                    bbox=table_bbox,
                    table_matrix=matrix,
                )
            )

        # Remove suppressed blocks from existing blocks
        clean_blocks = [
            b for idx, b in enumerate(existing_blocks)
            if idx not in suppressed_block_indices
        ]

        # Filter out drawings inside detected table regions so ImageHandler
        # doesn't mistake table grid lines for a flowchart / diagram
        remaining_drawings = [
            d for d in raw_drawings
            if not self._drawing_in_tables(d, table_bboxes)
        ]

        return clean_blocks + table_blocks, remaining_drawings

    # ------------------------------------------------------------------
    # Line segment extraction
    # ------------------------------------------------------------------

    def _extract_line_segments(
        self, drawings: list[dict]
    ) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
        """
        Extracts horizontal segments (y, x0, x1) and vertical segments (x, y0, y1)
        from drawing items and thin rects.
        """
        h_segments: list[tuple[float, float, float]] = []
        v_segments: list[tuple[float, float, float]] = []

        for d in drawings:
            # 1. Check item primitives in drawing
            for item in d.get("items", []):
                cmd = item[0]
                if cmd == "l":  # Line from p1 to p2
                    p1, p2 = item[1], item[2]
                    # Horizontal line
                    if abs(p1.y - p2.y) <= self.line_tolerance:
                        length = abs(p1.x - p2.x)
                        if length >= _MIN_LINE_LENGTH:
                            y = (p1.y + p2.y) / 2.0
                            h_segments.append((y, min(p1.x, p2.x), max(p1.x, p2.x)))
                    # Vertical line
                    elif abs(p1.x - p2.x) <= self.line_tolerance:
                        length = abs(p1.y - p2.y)
                        if length >= _MIN_LINE_LENGTH:
                            x = (p1.x + p2.x) / 2.0
                            v_segments.append((x, min(p1.y, p2.y), max(p1.y, p2.y)))

                elif cmd == "re":  # Rectangle
                    r = fitz.Rect(item[1])
                    if r.height <= self.line_tolerance and r.width >= _MIN_LINE_LENGTH:
                        y = (r.y0 + r.y1) / 2.0
                        h_segments.append((y, r.x0, r.x1))
                    elif r.width <= self.line_tolerance and r.height >= _MIN_LINE_LENGTH:
                        x = (r.x0 + r.x1) / 2.0
                        v_segments.append((x, r.y0, r.y1))

            # 2. Check outer rect if no items or items missed
            rect = d.get("rect")
            if rect:
                r = fitz.Rect(rect)
                if r.height <= self.line_tolerance and r.width >= _MIN_LINE_LENGTH:
                    y = (r.y0 + r.y1) / 2.0
                    h_segments.append((y, r.x0, r.x1))
                elif r.width <= self.line_tolerance and r.height >= _MIN_LINE_LENGTH:
                    x = (r.x0 + r.x1) / 2.0
                    v_segments.append((x, r.y0, r.y1))

        return h_segments, v_segments

    # ------------------------------------------------------------------
    # Grid construction & clustering
    # ------------------------------------------------------------------

    def _cluster_grids(
        self,
        h_segments: list[tuple[float, float, float]],
        v_segments: list[tuple[float, float, float]],
        page_num: int,
    ) -> list[dict]:
        """
        Groups intersecting horizontal and vertical lines into distinct
        table bounding boxes and row/column coordinate lists.
        """
        active_h = list(h_segments)
        active_v = list(v_segments)

        groups: list[dict] = []

        for h in active_h:
            hy, hx0, hx1 = h
            # Find all v lines that intersect or touch this h line
            matching_v = [
                v for v in active_v
                if (hx0 - self.line_tolerance) <= v[0] <= (hx1 + self.line_tolerance)
                and (v[1] - self.line_tolerance) <= hy <= (v[2] + self.line_tolerance)
            ]
            if not matching_v:
                continue

            # Check if this intersects an existing group
            merged = False
            for grp in groups:
                if (
                    hy >= (grp["y0"] - 10.0)
                    and hy <= (grp["y1"] + 10.0)
                    and not (hx1 < grp["x0"] - 10.0 or hx0 > grp["x1"] + 10.0)
                ):
                    grp["h_lines"].append(hy)
                    grp["x0"] = min(grp["x0"], hx0)
                    grp["x1"] = max(grp["x1"], hx1)
                    grp["y0"] = min(grp["y0"], hy)
                    grp["y1"] = max(grp["y1"], hy)
                    for mv in matching_v:
                        grp["v_lines"].append(mv[0])
                        grp["x0"] = min(grp["x0"], mv[0])
                        grp["x1"] = max(grp["x1"], mv[0])
                        grp["y0"] = min(grp["y0"], mv[1])
                        grp["y1"] = max(grp["y1"], mv[2])
                    merged = True
                    break

            if not merged:
                v_xs = [mv[0] for mv in matching_v]
                v_y0s = [mv[1] for mv in matching_v]
                v_y1s = [mv[2] for mv in matching_v]
                groups.append({
                    "x0": min([hx0] + v_xs),
                    "x1": max([hx1] + v_xs),
                    "y0": min([hy] + v_y0s),
                    "y1": max([hy] + v_y1s),
                    "h_lines": [hy],
                    "v_lines": v_xs,
                })

        # Process each group into a table grid
        tables: list[dict] = []
        for grp in groups:
            unique_h = self._cluster_coordinates(grp["h_lines"])
            unique_v = self._cluster_coordinates(grp["v_lines"])

            if len(unique_h) < 2 or len(unique_v) < 2:
                continue

            x0, x1 = unique_v[0], unique_v[-1]
            y0, y1 = unique_h[0], unique_h[-1]

            width = x1 - x0
            height = y1 - y0

            if width < _MIN_TABLE_WIDTH or height < _MIN_TABLE_HEIGHT:
                continue

            tables.append({
                "bbox": BoundingBox(x0=x0, y0=y0, x1=x1, y1=y1, page_num=page_num),
                "rows": unique_h,
                "cols": unique_v,
            })

        return tables

    def _cluster_coordinates(self, coords: list[float]) -> list[float]:
        if not coords:
            return []
        s = sorted(coords)
        clustered = [s[0]]
        for c in s[1:]:
            if abs(c - clustered[-1]) > self.line_tolerance:
                clustered.append(c)
        return clustered

    # ------------------------------------------------------------------
    # Geometry helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _is_cell_content(inner: BoundingBox, cell: BoundingBox) -> bool:
        """
        True when the text block belongs inside the cell.
        Uses center point with a small tolerance so text touching grid borders
        is correctly assigned.
        """
        cx = inner.center_x
        cy = inner.center_y
        pad = 2.0
        return (
            (cell.x0 - pad) <= cx <= (cell.x1 + pad)
            and (cell.y0 - pad) <= cy <= (cell.y1 + pad)
        )

    @staticmethod
    def _drawing_in_tables(d: dict, table_bboxes: list[BoundingBox]) -> bool:
        """True if the drawing primitive lies completely inside a table bbox."""
        rect = d.get("rect")
        if not rect:
            return False
        r = fitz.Rect(rect)
        for t in table_bboxes:
            if (
                r.x0 >= (t.x0 - 4.0)
                and r.y0 >= (t.y0 - 4.0)
                and r.x1 <= (t.x1 + 4.0)
                and r.y1 <= (t.y1 + 4.0)
            ):
                return True
        return False
