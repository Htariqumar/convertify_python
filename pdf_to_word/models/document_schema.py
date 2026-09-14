"""
Foundation data contract for the entire PDF-to-Word pipeline.

Every stage — analyzer, classifiers, handlers, builder — communicates
exclusively through these dataclasses.  Nothing in this file imports
from any other pipeline module, so it can be safely imported anywhere
without circular-dependency risk.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


# ---------------------------------------------------------------------------
# Element Types
# ---------------------------------------------------------------------------

class ElementType(Enum):
    PARAGRAPH    = "paragraph"
    HEADING      = "heading"
    LIST_ITEM    = "list_item"
    TABLE        = "table"
    IMAGE        = "image"
    VECTOR_DIAGRAM = "vector_diagram"
    HEADER       = "header"   # running page header
    FOOTER       = "footer"   # running page footer


# ---------------------------------------------------------------------------
# Geometry
# ---------------------------------------------------------------------------

@dataclass
class BoundingBox:
    """
    PDF coordinate system: origin is top-left, units are points (pt).
    page_num is 0-indexed (matches fitz.Page.number).
    """
    x0: float
    y0: float
    x1: float
    y1: float
    page_num: int = 0

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2.0

    def overlaps(self, other: "BoundingBox", threshold: float = 0.5) -> bool:
        """True when this box's intersection with *other* covers >= threshold
        of this box's own area (avoids false positives from thin slivers)."""
        ix0 = max(self.x0, other.x0)
        iy0 = max(self.y0, other.y0)
        ix1 = min(self.x1, other.x1)
        iy1 = min(self.y1, other.y1)
        if ix1 <= ix0 or iy1 <= iy0:
            return False
        inter = (ix1 - ix0) * (iy1 - iy0)
        own_area = max(1.0, self.width * self.height)
        return (inter / own_area) >= threshold


# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------

@dataclass
class TextSpan:
    """
    Smallest styled text unit — maps directly to a python-docx Run.
    color_hex is always lowercase 6-digit hex WITHOUT the '#' prefix
    (e.g. '000000' for black) so downstream code can do a direct
    int(color_hex, 16) without stripping.
    """
    text: str
    font_name: str  = "Arial"
    font_size: float = 11.0
    is_bold: bool   = False
    is_italic: bool = False
    color_hex: str  = "000000"   # no '#' prefix — see docstring
    hyperlink: Optional[str] = None


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------

@dataclass
class TableCellData:
    bbox:     BoundingBox
    spans:    List[TextSpan] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.spans)


# ---------------------------------------------------------------------------
# Content Block
# ---------------------------------------------------------------------------

@dataclass
class Block:
    """
    A single logical content unit on a page (paragraph, heading, table, …).

    Depending on element_type, only certain optional fields are populated:
      • PARAGRAPH / HEADING / LIST_ITEM / HEADER / FOOTER → spans
      • HEADING                                           → heading_level
      • LIST_ITEM                                         → list_type
      • TABLE                                             → table_matrix
      • IMAGE / VECTOR_DIAGRAM                            → image_bytes, image_format
    """
    element_type: ElementType
    bbox:         BoundingBox
    order_index:  int = 0

    # Text content (PARAGRAPH, HEADING, LIST_ITEM, HEADER, FOOTER)
    spans: List[TextSpan] = field(default_factory=list)

    # Heading metadata
    heading_level: Optional[int] = None   # 1 | 2 | 3

    # List metadata
    list_type: Optional[str] = None       # 'bullet' | 'numbered'

    # Table content
    table_matrix: List[List[TableCellData]] = field(default_factory=list)

    # Image / diagram content
    image_bytes:  Optional[bytes] = None
    image_format: str = "png"

    @property
    def plain_text(self) -> str:
        """Convenience: all span text concatenated (no style info)."""
        return "".join(s.text for s in self.spans)


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

@dataclass
class PageGeometry:
    """
    All measurements are in PDF points (1 pt = 1/72 inch).
    python-docx's Pt() helper converts them to EMUs transparently.
    """
    page_num:      int
    width:         float
    height:        float
    margin_top:    float = 72.0   # 1 inch default
    margin_bottom: float = 72.0
    margin_left:   float = 72.0
    margin_right:  float = 72.0

    @property
    def is_landscape(self) -> bool:
        return self.width > self.height

    @property
    def usable_width(self) -> float:
        return max(0.0, self.width - self.margin_left - self.margin_right)

    @property
    def usable_height(self) -> float:
        return max(0.0, self.height - self.margin_top - self.margin_bottom)


@dataclass
class DocumentPage:
    geometry: PageGeometry
    blocks:   List[Block] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Document
# ---------------------------------------------------------------------------

@dataclass
class DocumentModel:
    pages:  List[DocumentPage] = field(default_factory=list)
    title:  Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
