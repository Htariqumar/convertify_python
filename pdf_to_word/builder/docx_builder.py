"""
Word document builder — assembles a python-docx Document from a DocumentModel.

Design decisions & bug fixes applied:

  Bug 3 fix — section handling:
    A new Word section (docx.add_section()) is inserted ONLY when the
    page dimensions actually change (Portrait ↔ Landscape or a real size
    change).  Same-size consecutive pages get a simple page break instead,
    which avoids creating per-page header/footer zones and keeps the
    document editable as a whole.

  Header / Footer:
    Blocks classified as HEADER or FOOTER by ElementClassifier are written
    into Word's real header/footer area (section.header / section.footer)
    rather than the document body, so they behave like running headers in
    Word (show on every page, don't reflow with body text).

  Hyperlinks:
    External URI hyperlinks stored in TextSpan.hyperlink are converted to
    genuine Word hyperlinks via the OxmlElement / r:id mechanism rather
    than just underlined text.

  Color:
    TextSpan.color_hex is a 6-digit hex string WITHOUT '#'.  RGBColor()
    accepts three int values, which we derive with a single int(..., 16)
    call — no stripping or slicing gymnastics.
"""
from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from ..models.document_schema import (
    Block,
    DocumentModel,
    ElementType,
    PageGeometry,
    TextSpan,
)

# Tolerance (in points) for "same page size" check.
_SIZE_TOLERANCE: float = 1.0


class DocxBuilder:
    def __init__(self, doc_model: DocumentModel) -> None:
        self.doc_model   = doc_model
        self.docx        = Document()
        # Remove the empty paragraph that python-docx always starts with
        self._remove_default_paragraph()

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def build(self, output_path: str) -> None:
        prev_geo: PageGeometry | None = None
        first_header_written = False
        first_footer_written = False

        for p_idx, page in enumerate(self.doc_model.pages):
            geo = page.geometry

            # --- Section / page-break logic (Bug 3 fix) ---
            if p_idx == 0:
                section = self.docx.sections[0]
                self._apply_geometry(section, geo)
            else:
                assert prev_geo is not None
                if self._size_changed(prev_geo, geo):
                    section = self.docx.add_section()
                    self._apply_geometry(section, geo)
                else:
                    self.docx.add_page_break()
                    section = self.docx.sections[-1]

            # --- Collect headers / footers for this page ---
            header_blocks = [b for b in page.blocks if b.element_type == ElementType.HEADER]
            footer_blocks = [b for b in page.blocks if b.element_type == ElementType.FOOTER]
            body_blocks   = [b for b in page.blocks if b.element_type not in (
                ElementType.HEADER, ElementType.FOOTER
            )]

            # Write headers/footers into Word's real header/footer (once
            # per unique section — repeated text is handled automatically
            # by Word's "Link to Previous" behaviour).
            if header_blocks and not first_header_written:
                self._write_header(section, header_blocks)
                first_header_written = True

            if footer_blocks and not first_footer_written:
                self._write_footer(section, footer_blocks)
                first_footer_written = True

            # --- Body blocks ---
            for block in body_blocks:
                self._render_block(block)

            prev_geo = geo

        # Document metadata
        core = self.docx.core_properties
        if self.doc_model.title:
            core.title = self.doc_model.title
        if self.doc_model.author:
            core.author = self.doc_model.author

        self.docx.save(output_path)

    # ------------------------------------------------------------------
    # Section helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_geometry(section, geo: PageGeometry) -> None:
        section.page_width    = Pt(geo.width)
        section.page_height   = Pt(geo.height)
        section.top_margin    = Pt(geo.margin_top)
        section.bottom_margin = Pt(geo.margin_bottom)
        section.left_margin   = Pt(geo.margin_left)
        section.right_margin  = Pt(geo.margin_right)

    @staticmethod
    def _size_changed(a: PageGeometry, b: PageGeometry) -> bool:
        return (
            abs(a.width  - b.width)  > _SIZE_TOLERANCE
            or abs(a.height - b.height) > _SIZE_TOLERANCE
        )

    # ------------------------------------------------------------------
    # Header / Footer
    # ------------------------------------------------------------------

    def _write_header(self, section, blocks: list[Block]) -> None:
        hdr = section.header
        # Clear the default empty paragraph
        for p in hdr.paragraphs:
            p.clear()
        para = hdr.paragraphs[0]
        for block in blocks:
            self._add_spans_to_paragraph(para, block.spans)

    def _write_footer(self, section, blocks: list[Block]) -> None:
        ftr = section.footer
        for p in ftr.paragraphs:
            p.clear()
        para = ftr.paragraphs[0]
        for block in blocks:
            self._add_spans_to_paragraph(para, block.spans)

    # ------------------------------------------------------------------
    # Block rendering
    # ------------------------------------------------------------------

    def _render_block(self, block: Block) -> None:
        if block.element_type == ElementType.HEADING:
            self._render_heading(block)
        elif block.element_type in (ElementType.PARAGRAPH, ElementType.LIST_ITEM):
            self._render_paragraph(block)
        elif block.element_type == ElementType.TABLE:
            self._render_table(block)
        elif block.element_type in (ElementType.IMAGE, ElementType.VECTOR_DIAGRAM):
            self._render_image(block)
        # HEADER / FOOTER are handled before this loop — no else needed

    def _render_heading(self, block: Block) -> None:
        level = block.heading_level or 1
        para  = self.docx.add_heading(level=level)
        para.paragraph_format.space_after = Pt(4)
        self._add_spans_to_paragraph(para, block.spans)

    def _render_paragraph(self, block: Block) -> None:
        para = self.docx.add_paragraph()
        para.paragraph_format.space_after = Pt(2)

        if block.element_type == ElementType.LIST_ITEM:
            style = (
                "List Bullet"   if block.list_type == "bullet"
                else "List Number"
            )
            try:
                para.style = self.docx.styles[style]
            except KeyError:
                pass   # style not in this template → plain paragraph

        self._add_spans_to_paragraph(para, block.spans)

    def _render_table(self, block: Block) -> None:
        matrix = block.table_matrix
        if not matrix:
            return
        rows = len(matrix)
        cols = max(len(r) for r in matrix)
        if rows == 0 or cols == 0:
            return

        try:
            tbl = self.docx.add_table(rows=rows, cols=cols)
            tbl.style = "Table Grid"

            for r_idx, row in enumerate(matrix):
                for c_idx in range(cols):
                    if c_idx >= len(row) or row[c_idx] is None:
                        continue
                    cell_data = row[c_idx]
                    cell = tbl.cell(r_idx, c_idx)
                    p = cell.paragraphs[0]
                    p.clear()
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)

                    if cell_data.spans:
                        self._add_spans_to_paragraph(p, cell_data.spans)
                    elif cell_data.text:
                        run = p.add_run(cell_data.text)
                        run.font.name = "Arial"
                        run.font.size = Pt(10)
        except Exception:
            pass  # Never allow table formatting glitch to abort full document assembly

    def _render_image(self, block: Block) -> None:
        if not block.image_bytes:
            return
        from io import BytesIO
        try:
            # Cap width at the usable page width of the first section
            max_width_pt = self.docx.sections[0].page_width.pt - (
                self.docx.sections[0].left_margin.pt
                + self.docx.sections[0].right_margin.pt
            )
            img_width_pt = block.bbox.width
            width_in     = min(img_width_pt, max_width_pt) / 72.0
            self.docx.add_picture(BytesIO(block.image_bytes), width=Inches(width_in))
        except Exception:
            pass   # Never let a bad image kill the whole build

    # ------------------------------------------------------------------
    # Span → Run rendering
    # ------------------------------------------------------------------

    def _add_spans_to_paragraph(self, para, spans: list[TextSpan]) -> None:
        for span in spans:
            if span.hyperlink:
                self._add_hyperlink_run(para, span)
            else:
                run = para.add_run(span.text)
                self._apply_span_style(run, span)

    @staticmethod
    def _apply_span_style(run, span: TextSpan) -> None:
        run.bold   = span.is_bold
        run.italic = span.is_italic
        if span.font_name:
            run.font.name = span.font_name
        if span.font_size > 0:
            run.font.size = Pt(span.font_size)
        try:
            color_int = int(span.color_hex, 16)
            r = (color_int >> 16) & 0xFF
            g = (color_int >>  8) & 0xFF
            b =  color_int        & 0xFF
            # Only set non-black colors to avoid overriding heading styles
            if (r, g, b) != (0, 0, 0):
                run.font.color.rgb = RGBColor(r, g, b)
        except (ValueError, TypeError):
            pass

    def _add_hyperlink_run(self, para, span: TextSpan) -> None:
        """
        Adds a genuine Word hyperlink (not just underlined blue text).
        The relationship is added to the paragraph's part, and the run
        is wrapped in a w:hyperlink element referencing that relationship.
        """
        try:
            part = para.part
            r_id = part.relate_to(
                span.hyperlink,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            hyperlink_el = OxmlElement("w:hyperlink")
            hyperlink_el.set(qn("r:id"), r_id)

            new_run = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")

            # Style: underline + blue (standard hyperlink appearance)
            u_el = OxmlElement("w:u")
            u_el.set(qn("w:val"), "single")
            color_el = OxmlElement("w:color")
            color_el.set(qn("w:val"), "0563C1")

            rPr.append(u_el)
            rPr.append(color_el)
            new_run.append(rPr)

            t_el = OxmlElement("w:t")
            t_el.text = span.text
            new_run.append(t_el)
            hyperlink_el.append(new_run)
            para._p.append(hyperlink_el)
        except Exception:
            # Fallback to plain run if hyperlink wiring fails
            run = para.add_run(span.text)
            self._apply_span_style(run, span)

    # ------------------------------------------------------------------
    # Init helper
    # ------------------------------------------------------------------

    def _remove_default_paragraph(self) -> None:
        """python-docx always adds one empty paragraph on Document() —
        remove it so the first real block doesn't have a blank line above it."""
        body = self.docx.element.body
        # Only remove if it is truly empty
        paras = body.findall(qn("w:p"))
        if paras and not paras[0].text:
            body.remove(paras[0])
