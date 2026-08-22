from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pdf2docx import Converter
import pdfplumber
from openpyxl import Workbook
from pptx import Presentation
import fitz # PyMuPDF
import tempfile
import os
import re
import shutil
import subprocess
import zipfile
import edge_tts
from pydantic import BaseModel
import imageio_ffmpeg

# LibreOffice (word/excel/ppt <-> pdf) and Ghostscript (compress-pdf, thumbnails) run here
# rather than in the Vercel Next.js app because neither fits a serverless function: Vercel
# caps deployment size at 250MB and LibreOffice alone is 300MB-1GB+, so it can never be
# bundled there regardless of install method. This host is a normal Linux container/VM, so
# both are installed as regular system packages - see the Dockerfile.
SOFFICE_BIN = os.environ.get("SOFFICE_PATH", "soffice")
GHOSTSCRIPT_BIN = os.environ.get("GHOSTSCRIPT_PATH", "gs")

# Automatically append the bundled ffmpeg to the system PATH
os.environ["PATH"] += os.pathsep + os.path.dirname(imageio_ffmpeg.get_ffmpeg_exe())

import faster_whisper
# Load model once at startup to avoid loading on every request
whisper_model = faster_whisper.WhisperModel("base", device="cpu", compute_type="int8")

class TTSRequest(BaseModel):
    text: str
    voice: str = "en-US-JennyNeural"
    speed: str = "+0%"

app = FastAPI(title="Convertify Python Microservice")

@app.get("/")
def read_root():
    return {"message": "Python Microservice is running!"}

def remove_files(paths):
    for path in paths:
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            print(f"Error removing {path}: {e}")


def convert_with_soffice(input_path: str, output_dir: str, target_format: str = "pdf") -> str:
    """Converts a document to the given format via headless LibreOffice (e.g. word/excel/ppt -> pdf,
    or pdf -> docx). Each call gets its own -env:UserInstallation profile dir so concurrent requests
    don't collide on LibreOffice's single-instance profile lock."""
    profile_dir = tempfile.mkdtemp()
    try:
        result = subprocess.run(
            [
                SOFFICE_BIN, "--headless", "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", target_format, "--outdir", output_dir, input_path,
            ],
            capture_output=True,
            timeout=100,
        )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.{target_format}")
    if result.returncode != 0 or not os.path.exists(output_path):
        stderr = result.stderr.decode(errors="ignore") if result.stderr else "unknown error"
        raise RuntimeError(f"LibreOffice conversion failed: {stderr}")
    return output_path


def prepare_excel_for_pdf(input_path: str, ext: str) -> None:
    """Force Calc's 'fit sheet to page width' plus a tight print area before handing off
    to LibreOffice. Headless --convert-to otherwise leaves pagination entirely to whatever
    the workbook happened to save - for most files exported from a web app or built from
    scraped/pasted data that's nothing, so a wide sheet gets sliced across several PDF
    pages mid-table instead of shrinking to fit one page wide, the way Excel's own
    print preview normally behaves."""
    if ext not in (".xlsx", ".xlsm"):
        return
    try:
        from openpyxl import load_workbook
        wb = load_workbook(input_path)
        for ws in wb.worksheets:
            if ws.max_row < 1 or ws.max_column < 1:
                continue
            ws.print_area = ws.dimensions
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            if ws.max_column > 8:
                ws.page_setup.orientation = "landscape"
        wb.save(input_path)
    except Exception as e:
        # Best-effort only - if the workbook can't be parsed (password-protected,
        # unusual dialect), fall back to LibreOffice's default conversion rather than
        # failing the whole request.
        print(f"[excel-to-pdf] print-setup pre-processing skipped: {e}")


async def office_to_pdf_response(background_tasks: BackgroundTasks, file: UploadFile) -> FileResponse:
    """Shared handler for word/excel/ppt -> pdf: saves the upload, converts via LibreOffice,
    and returns the resulting PDF. Keeps the original extension so LibreOffice picks the
    right import filter (works for both legacy .doc/.xls/.ppt and modern .docx/.xlsx/.pptx)."""
    ext = os.path.splitext(file.filename or "")[1] or ".tmp"
    work_dir = tempfile.mkdtemp()
    input_path = os.path.join(work_dir, f"input{ext}")

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        prepare_excel_for_pdf(input_path, ext)

        output_path = convert_with_soffice(input_path, work_dir)

        background_tasks.add_task(shutil.rmtree, work_dir, True)
        return FileResponse(path=output_path, filename=f"converted_{file.filename}.pdf", media_type="application/pdf")
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Create temporary files for the input PDF and output DOCX
    fd_pdf, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
    fd_docx, temp_docx_path = tempfile.mkstemp(suffix=".docx")
    
    os.close(fd_pdf)
    os.close(fd_docx)

    try:
        # Write the uploaded file to the temp PDF path
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        from pdf2docx import Converter
        import docx
        import fitz
        
        def has_spacing_issue(docx_path):
            try:
                doc = docx.Document(docx_path)
                long_words = 0
                total_words = 0
                for para in doc.paragraphs:
                    words = para.text.split()
                    for w in words:
                        total_words += 1
                        if len(w) >= 15:
                            long_words += 1
                if total_words == 0: return False
                return (long_words / total_words) > 0.05
            except:
                return False

        def fix_glued_words_in_docx(docx_path):
            # pdf2docx occasionally drops the space between words (common on
            # Canva-exported PDFs, justified text, etc.), which has_spacing_issue()
            # detects. Rather than throwing away pdf2docx's layout (tables, images,
            # columns, fonts - all correct) and rebuilding the page from scratch,
            # try a much cheaper surgical fix first: re-segment only the glued
            # tokens in place with a dictionary-based word splitter, leaving every
            # other run and all formatting untouched. Falls back to the full
            # from-scratch rebuild only if this doesn't clear the issue.
            import wordninja

            def split_token(token):
                if not token.isalpha() or len(token) < 15:
                    return token
                parts = wordninja.split(token)
                if len(parts) <= 1 or ''.join(parts) != token:
                    return token
                return ' '.join(parts)

            def fix_paragraphs(paragraphs):
                for para in paragraphs:
                    for run in para.runs:
                        if not run.text:
                            continue
                        words = run.text.split(' ')
                        fixed = [split_token(w) for w in words]
                        if fixed != words:
                            run.text = ' '.join(fixed)

            try:
                doc = docx.Document(docx_path)
                fix_paragraphs(doc.paragraphs)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            fix_paragraphs(cell.paragraphs)
                doc.save(docx_path)
                return not has_spacing_issue(docx_path)
            except:
                return False

        def convert_with_correct_spacing(pdf_path, docx_path):
            import re
            from io import BytesIO
            from collections import Counter
            from docx.shared import Pt, Inches
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement

            pdf = fitz.open(pdf_path)
            out_doc = docx.Document()
            section = out_doc.sections[0]
            usable_width_in = section.page_width.inches - section.left_margin.inches - section.right_margin.inches

            def clean_font_name(f_str):
                name = f_str.split('+')[-1]
                name = re.sub(r'-(Bold|Italic|Regular|Medium|SemiBold|Light).*', '', name, flags=re.IGNORECASE)
                name = re.sub(r'(MT|PSMT|MS|PS)$', '', name, flags=re.IGNORECASE)
                name = re.sub(r'([a-z])([A-Z])', r'\1 \2', name)
                return name.strip()

            def add_page_number_field(paragraph):
                run = paragraph.add_run()
                begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
                instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = "PAGE"
                end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
                run._r.append(begin); run._r.append(instr); run._r.append(end)

            def apply_field_text(paragraph, sample):
                m = re.search(r'\d+', sample)
                if not m:
                    paragraph.add_run(sample)
                    return
                paragraph.add_run(sample[:m.start()])
                add_page_number_field(paragraph)
                paragraph.add_run(sample[m.end():])

            def overlap_ratio(a, b):
                ax0, ay0, ax1, ay1 = a
                bx0, by0, bx1, by1 = b
                ix0, iy0 = max(ax0, bx0), max(ay0, by0)
                ix1, iy1 = min(ax1, bx1), min(ay1, by1)
                if ix1 <= ix0 or iy1 <= iy0:
                    return 0.0
                inter = (ix1 - ix0) * (iy1 - iy0)
                area_a = max(1.0, (ax1 - ax0) * (ay1 - ay0))
                return inter / area_a

            def detect_columns(items, page_x0, page_x1):
                # Finds vertical whitespace gutters shared by most lines, so multi-column
                # PDFs (magazines/papers) read top-to-bottom per column instead of
                # interleaving alternating lines from the left and right columns.
                page_width = page_x1 - page_x0
                if page_width <= 0 or len(items) < 6:
                    return [(page_x0, page_x1)]
                bins = 80
                bin_w = page_width / bins
                covered = [False] * bins
                for it in items:
                    x0, x1 = it["bbox"][0], it["bbox"][2]
                    b0 = max(0, int((x0 - page_x0) / bin_w))
                    b1 = min(bins - 1, int((x1 - page_x0) / bin_w))
                    for i in range(b0, b1 + 1):
                        covered[i] = True
                gaps = []
                i = 0
                while i < bins:
                    if not covered[i]:
                        j = i
                        while j < bins and not covered[j]:
                            j += 1
                        gaps.append((page_x0 + i * bin_w, page_x0 + j * bin_w, i, j - 1))
                        i = j
                    else:
                        i += 1
                real_gaps = [g for g in gaps if (g[1] - g[0]) >= 12 and g[2] > 3 and g[3] < bins - 4]
                if not real_gaps or len(real_gaps) > 2:
                    return [(page_x0, page_x1)]
                splits = sorted((g[0] + g[1]) / 2 for g in real_gaps)
                bands, prev = [], page_x0
                for s in splits:
                    bands.append((prev, s))
                    prev = s
                bands.append((prev, page_x1))
                counts = [sum(1 for it in items if bx0 <= (it["bbox"][0] + it["bbox"][2]) / 2 < bx1) for bx0, bx1 in bands]
                total = sum(counts)
                if total == 0 or min(counts) < max(2, total * 0.12):
                    return [(page_x0, page_x1)]
                return bands

            def render_table(item, state):
                rows = item["rows"]
                max_cols = max(len(r) for r in rows)
                table = out_doc.add_table(rows=len(rows), cols=max_cols)
                table.style = "Table Grid"
                for r_idx, row in enumerate(rows):
                    for c_idx in range(max_cols):
                        val = row[c_idx] if c_idx < len(row) and row[c_idx] is not None else ""
                        table.cell(r_idx, c_idx).text = str(val).strip()
                state["p"] = None
                state["last_y_bottom"] = -999
                state["last_font_size"] = -1

            def render_image(item, state):
                bbox = item["bbox"]
                width_in = min((bbox[2] - bbox[0]) / 72.0, usable_width_in)
                if width_in <= 0:
                    width_in = usable_width_in
                try:
                    out_doc.add_picture(BytesIO(item["data"]), width=Inches(width_in))
                except Exception:
                    pass
                state["p"] = None
                state["last_y_bottom"] = -999
                state["last_font_size"] = -1

            def render_text(item, state):
                bbox = item["bbox"]
                y_top, y_bottom = bbox[1], bbox[3]
                font_size = item["font_size"]
                vertical_gap = y_top - state["last_y_bottom"]
                is_bullet = item["full_text"].startswith(('•', '◦', '-', '*'))

                start_new_para = True
                if state["p"] is not None:
                    if is_bullet:
                        start_new_para = True
                    elif y_top < (state["last_y_bottom"] - font_size * 0.3):
                        start_new_para = True
                    elif vertical_gap < (font_size * 0.5) and abs(font_size - state["last_font_size"]) < 1.0:
                        start_new_para = False

                if start_new_para:
                    state["p"] = out_doc.add_paragraph()
                else:
                    state["p"].add_run(" ")

                for r in item["runs"]:
                    run = state["p"].add_run(r["text"])
                    raw_font = r["font"]
                    flags = r["flags"]
                    if (flags & 16) or "bold" in raw_font or "black" in raw_font:
                        run.bold = True
                    if (flags & 2) or "italic" in raw_font:
                        run.italic = True
                    clean_name = clean_font_name(raw_font)
                    if clean_name and clean_name.lower() != "unknown":
                        run.font.name = clean_name
                    if r["size"] > 0:
                        run.font.size = Pt(round(r["size"]))

                state["last_y_bottom"] = y_bottom
                state["last_font_size"] = font_size

            def render_band(band_items, page_x0, page_x1, state):
                if not band_items:
                    return
                col_bands = detect_columns(band_items, page_x0, page_x1)
                if len(col_bands) <= 1:
                    ordered = sorted(band_items, key=lambda x: (round(x["bbox"][1] / 5), x["bbox"][0]))
                    for it in ordered:
                        (render_image if it["type"] == "image" else render_text)(it, state)
                    return
                for bx0, bx1 in col_bands:
                    col_items = [it for it in band_items if bx0 <= (it["bbox"][0] + it["bbox"][2]) / 2 < bx1]
                    col_items.sort(key=lambda x: (round(x["bbox"][1] / 5), x["bbox"][0]))
                    for it in col_items:
                        (render_image if it["type"] == "image" else render_text)(it, state)

            page_count = len(pdf)

            # ---- Pass 1: detect header/footer lines that repeat across most pages, so
            # they land in Word's real header/footer instead of the document body. ----
            top_counter, bottom_counter = Counter(), Counter()
            top_samples, bottom_samples = {}, {}
            for page in pdf:
                h = page.rect.height
                for b in page.get_text("dict")["blocks"]:
                    if b.get("type") != 0:
                        continue
                    for l in b.get("lines", []):
                        txt = "".join(s.get("text", "") for s in l.get("spans", [])).strip()
                        if not txt:
                            continue
                        y0, y1 = l["bbox"][1], l["bbox"][3]
                        norm = re.sub(r'\d+', '#', txt)
                        if y1 < h * 0.10:
                            top_counter[norm] += 1
                            top_samples.setdefault(norm, txt)
                        elif y0 > h * 0.90:
                            bottom_counter[norm] += 1
                            bottom_samples.setdefault(norm, txt)

            repeat_threshold = max(2, int(page_count * 0.6))
            header_patterns = {p for p, c in top_counter.items() if c >= repeat_threshold} if page_count >= 3 else set()
            footer_patterns = {p for p, c in bottom_counter.items() if c >= repeat_threshold} if page_count >= 3 else set()

            if header_patterns:
                apply_field_text(section.header.paragraphs[0], top_samples[next(iter(header_patterns))])
            if footer_patterns:
                apply_field_text(section.footer.paragraphs[0], bottom_samples[next(iter(footer_patterns))])

            # ---- Pass 2: rebuild the body page by page (real page breaks, tables,
            # inline images and column-aware reading order) ----
            for page_index, page in enumerate(pdf):
                if page_index > 0:
                    out_doc.add_page_break()

                h = page.rect.height
                page_x0, page_x1 = page.rect.x0, page.rect.x1

                table_items = []
                try:
                    for t in page.find_tables().tables:
                        try:
                            rows = t.extract()
                        except Exception:
                            continue
                        if rows:
                            table_items.append({"type": "table", "bbox": list(t.bbox), "rows": rows})
                except Exception:
                    pass

                image_items = []
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    try:
                        rects = page.get_image_rects(xref) or []
                    except Exception:
                        rects = []
                    if not rects:
                        continue
                    try:
                        pix = fitz.Pixmap(pdf, xref)
                        if pix.n - pix.alpha >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        img_bytes = pix.tobytes("png")
                    except Exception:
                        continue
                    for rect in rects:
                        rbbox = [rect[0], rect[1], rect[2], rect[3]]
                        if (rbbox[2] - rbbox[0]) < 5 or (rbbox[3] - rbbox[1]) < 5:
                            continue
                        if any(overlap_ratio(rbbox, t["bbox"]) > 0.5 for t in table_items):
                            continue
                        image_items.append({"type": "image", "bbox": rbbox, "data": img_bytes})

                text_items = []
                for b in page.get_text("dict")["blocks"]:
                    if b.get("type") != 0:
                        continue
                    for l in b.get("lines", []):
                        spans = l.get("spans", [])
                        if not spans:
                            continue
                        bbox = l.get("bbox", [0, 0, 0, 0])
                        raw_text = "".join(s.get("text", "") for s in spans).strip()
                        if not raw_text:
                            continue
                        norm = re.sub(r'\d+', '#', raw_text)
                        in_header_zone = bbox[3] < h * 0.10
                        in_footer_zone = bbox[1] > h * 0.90
                        if (in_header_zone and norm in header_patterns) or (in_footer_zone and norm in footer_patterns):
                            continue
                        if any(overlap_ratio(bbox, t["bbox"]) > 0.5 for t in table_items):
                            continue

                        run_params = []
                        last_x = None
                        first_span = spans[0]
                        font_size = first_span.get("size", 11)

                        for s in spans:
                            text = s.get("text", "")
                            sbbox = s.get("bbox", [0, 0, 0, 0])

                            if last_x is not None:
                                gap = sbbox[0] - last_x
                                if gap > 2 and not text.startswith(" ") and not text.startswith(","):
                                    run_params.append({"text": " ", "size": font_size, "font": first_span.get("font", ""), "flags": 0})

                            run_params.append({"text": text, "size": s.get("size", 11), "font": s.get("font", "").lower(), "flags": s.get("flags", 0)})
                            last_x = sbbox[2]

                        full_text = "".join(r["text"] for r in run_params).strip()
                        if not full_text:
                            continue
                        text_items.append({"type": "text", "bbox": bbox, "runs": run_params, "full_text": full_text, "font_size": font_size})

                content_width = max((it["bbox"][2] for it in (text_items + image_items)), default=page_x1) - page_x0
                flowable = sorted(text_items + image_items + table_items, key=lambda it: (round(it["bbox"][1] / 5), it["bbox"][0]))

                state = {"p": None, "last_y_bottom": -999, "last_font_size": -1}
                band = []
                for it in flowable:
                    is_break = it["type"] == "table" or (it["type"] == "image" and (it["bbox"][2] - it["bbox"][0]) >= 0.7 * max(content_width, 1))
                    if is_break:
                        render_band(band, page_x0, page_x1, state)
                        band = []
                        (render_table if it["type"] == "table" else render_image)(it, state)
                    else:
                        band.append(it)
                render_band(band, page_x0, page_x1, state)

            out_doc.save(docx_path)
        
        # 1. Primary conversion using pdf2docx (fast, good enough for most straightforward PDFs)
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        used_method = "layout-mode"
        if has_spacing_issue(temp_docx_path):
            if fix_glued_words_in_docx(temp_docx_path):
                used_method = "layout-mode-corrected"
            else:
                convert_with_correct_spacing(temp_pdf_path, temp_docx_path)
                used_method = "text-mode"

        background_tasks.add_task(remove_files, [temp_pdf_path, temp_docx_path])
        return FileResponse(
            path=temp_docx_path, 
            filename=f"converted_{file.filename}.docx", 
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"X-Conversion-Method": used_method}
        )
    except Exception as e:
        remove_files([temp_pdf_path, temp_docx_path])
        raise HTTPException(status_code=500, detail=str(e))

def _pdf_cell_to_excel(raw):
    """Turn a scraped table cell into a native Excel type (+ number format) where it's
    unambiguously safe to do so, so numbers land as real numbers (sortable/summable,
    right-aligned) instead of importing as text - the single biggest visible gap between
    a scraped-looking sheet and one a person typed by hand. Returns (value, number_format)."""
    if raw is None:
        return None, None
    text = str(raw).strip()
    if text == "":
        return None, None
    # Leading-zero strings are almost always identifiers (zip codes, IDs), not numbers -
    # converting "00123" to 123 would silently corrupt the data.
    if re.fullmatch(r"0\d+", text):
        return text, None
    cleaned = text
    is_negative = cleaned.startswith("(") and cleaned.endswith(")")
    if is_negative:
        cleaned = cleaned[1:-1]
    cleaned = cleaned.replace(",", "")
    is_percent = cleaned.endswith("%")
    if is_percent:
        cleaned = cleaned[:-1]
    currency_match = re.match(r"^([$€£])", cleaned)
    if currency_match:
        cleaned = cleaned[1:]
    if not re.fullmatch(r"-?\d+(\.\d+)?", cleaned or ""):
        return text, None
    num = float(cleaned)
    if is_negative:
        num = -num
    if is_percent:
        return num / 100, "0.00%"
    if currency_match:
        return num, f'{currency_match.group(1)}#,##0.00'
    if num == int(num) and "." not in cleaned:
        return int(num), None
    return num, None


@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    fd_pdf, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
    fd_xlsx, temp_xlsx_path = tempfile.mkstemp(suffix=".xlsx")

    os.close(fd_pdf)
    os.close(fd_xlsx)

    try:
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        from openpyxl import Workbook
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
        import pdfplumber

        wb = Workbook()
        ws = wb.active
        header_font = Font(bold=True)
        max_col_widths = {}

        def write_row(row_idx, values, bold=False):
            for col_idx, raw in enumerate(values, start=1):
                value, number_format = _pdf_cell_to_excel(raw)
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if number_format:
                    cell.number_format = number_format
                if bold:
                    cell.font = header_font
                width = len(str(raw)) if raw is not None else 0
                if width > max_col_widths.get(col_idx, 0):
                    max_col_widths[col_idx] = width

        current_row = 1
        prev_num_cols = None
        prev_header = None

        with pdfplumber.open(temp_pdf_path) as pdf:
            for page in pdf.pages:
                tables = [t for t in page.extract_tables() if t]

                if not tables:
                    # Fallback for this page: no ruled table detected (e.g. a resume
                    # page or a plain-text page in an otherwise tabular report). Split
                    # on runs of 2+ spaces/tabs so whitespace-aligned columns still
                    # land in separate cells instead of one giant column. Scoped to
                    # the page (not the whole document) so a document that mixes
                    # tables and prose pages doesn't silently drop the prose ones.
                    # layout=True preserves the PDF's original horizontal gaps between
                    # words - plain extract_text() collapses any run of whitespace to
                    # a single space, which would erase the very column alignment this
                    # fallback relies on to split lines into cells.
                    text = page.extract_text(layout=True)
                    if text:
                        # layout mode represents the page's top/bottom margins as blank
                        # lines (it maps the full page height to a text grid) - trim
                        # those so the sheet doesn't open with a run of empty rows
                        # before the actual content. Blank lines *within* the content
                        # are kept since those are real paragraph/section breaks.
                        lines = text.split('\n')
                        while lines and not lines[0].strip():
                            lines.pop(0)
                        while lines and not lines[-1].strip():
                            lines.pop()

                        if lines and prev_num_cols is not None:
                            current_row += 1
                        for line in lines:
                            if not line.strip():
                                current_row += 1
                                continue
                            parts = re.split(r"\s{2,}|\t", line.strip())
                            write_row(current_row, parts)
                            current_row += 1
                        prev_num_cols = None
                        prev_header = None
                    continue

                for table in tables:
                    raw_rows = [["" if c is None else str(c).strip() for c in row] for row in table]
                    num_cols = len(raw_rows[0]) if raw_rows else 0

                    # A table that resumes on the next page re-extracts with the same
                    # column count and a repeated header row - treat it as a continuation
                    # of the previous block (no blank separator, no duplicate header)
                    # rather than a brand-new table.
                    is_continuation = (
                        prev_num_cols == num_cols
                        and prev_header is not None
                        and raw_rows
                        and raw_rows[0] == prev_header
                    )

                    rows_to_write = raw_rows[1:] if is_continuation else raw_rows
                    if not is_continuation and prev_num_cols is not None:
                        current_row += 1  # blank row between genuinely distinct tables

                    for i, row in enumerate(rows_to_write):
                        is_header_row = not is_continuation and i == 0
                        write_row(current_row, row, bold=is_header_row)
                        current_row += 1

                    prev_num_cols = num_cols
                    prev_header = raw_rows[0] if raw_rows else None

        for col_idx, width in max_col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width + 2, 8), 60)

        wb.save(temp_xlsx_path)

        background_tasks.add_task(remove_files, [temp_pdf_path, temp_xlsx_path])
        return FileResponse(path=temp_xlsx_path, filename=f"converted_{file.filename}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        remove_files([temp_pdf_path, temp_xlsx_path])
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/pdf-to-ppt")
async def convert_pdf_to_ppt(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    fd_pdf, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
    fd_pptx, temp_pptx_path = tempfile.mkstemp(suffix=".pptx")
    
    os.close(fd_pdf)
    os.close(fd_pptx)

    try:
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        doc = fitz.open(temp_pdf_path)
        prs = Presentation()
        from pptx.util import Pt
        
        if len(doc) > 0:
            # Set slide size to match the first PDF page
            prs.slide_width = Pt(doc[0].rect.width)
            prs.slide_height = Pt(doc[0].rect.height)
            
        blank_slide_layout = prs.slide_layouts[6] # blank layout
        
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img_data = pix.tobytes("png")
            
            fd_img, temp_img_path = tempfile.mkstemp(suffix=".png")
            os.close(fd_img)
            with open(temp_img_path, "wb") as f:
                f.write(img_data)
                
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(temp_img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            os.remove(temp_img_path)
            
        prs.save(temp_pptx_path)

        background_tasks.add_task(remove_files, [temp_pdf_path, temp_pptx_path])
        return FileResponse(path=temp_pptx_path, filename=f"converted_{file.filename}.pptx", media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        remove_files([temp_pdf_path, temp_pptx_path])
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/pdf-to-images")
async def convert_pdf_to_images(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    fd_pdf, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
    fd_zip, temp_zip_path = tempfile.mkstemp(suffix=".zip")
    
    os.close(fd_pdf)
    os.close(fd_zip)

    try:
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        doc = fitz.open(temp_pdf_path)
        
        with zipfile.ZipFile(temp_zip_path, 'w') as zipf:
            for i, page in enumerate(doc):
                pix = page.get_pixmap(dpi=150)
                img_data = pix.tobytes("jpeg")
                zipf.writestr(f"page_{i+1}.jpg", img_data)
                
        background_tasks.add_task(remove_files, [temp_pdf_path, temp_zip_path])
        return FileResponse(path=temp_zip_path, filename=f"converted_images.zip", media_type="application/zip")
    except Exception as e:
        remove_files([temp_pdf_path, temp_zip_path])
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/text-to-speech")
async def convert_text_to_speech(background_tasks: BackgroundTasks, request: TTSRequest):
    fd_mp3, temp_mp3_path = tempfile.mkstemp(suffix=".mp3")
    os.close(fd_mp3)

    try:
        # Edge-tts Communicate object
        communicate = edge_tts.Communicate(request.text, request.voice, rate=request.speed)
        
        # Generate and save audio to temp file
        await communicate.save(temp_mp3_path)
        
        background_tasks.add_task(remove_files, [temp_mp3_path])
        return FileResponse(path=temp_mp3_path, filename="speech.mp3", media_type="audio/mpeg")
    except Exception as e:
        remove_files([temp_mp3_path])
        print(f"[text-to-speech] edge-tts failed: {e}")
        # Must raise (not return a dict) so the response carries a non-200 status -
        # the Next.js caller requests this as arraybuffer and can't tell a JSON
        # error body from real audio bytes otherwise, so a 200 here silently
        # saves the error message as if it were a working MP3.
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/speech-to-text")
async def convert_speech_to_text(file: UploadFile = File(...), language: str = Form(None)):
    # Save the uploaded file temporarily
    fd_audio, temp_audio_path = tempfile.mkstemp(suffix=os.path.splitext(file.filename)[1])
    os.close(fd_audio)

    try:
        with open(temp_audio_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Auto-detection on the small "base" model can misfire for lower-resource
        # languages (e.g. Urdu getting misdetected as Chinese) - when the caller
        # knows the spoken language, passing it explicitly skips that guess and
        # is noticeably more accurate than relying on auto-detect.
        transcribe_kwargs = {"beam_size": 5}
        if language:
            transcribe_kwargs["language"] = language

        segments, info = whisper_model.transcribe(temp_audio_path, **transcribe_kwargs)

        # Combine segments into full text
        full_text = ""
        for segment in segments:
            full_text += segment.text + " "

        return {"text": full_text.strip(), "language": info.language}
    except Exception as e:
        return {"error": str(e)}
    finally:
        if os.path.exists(temp_audio_path):
            try:
                os.remove(temp_audio_path)
            except:
                pass

@app.post("/convert/word-to-pdf")
async def convert_word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    return await office_to_pdf_response(background_tasks, file)


@app.post("/convert/excel-to-pdf")
async def convert_excel_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    return await office_to_pdf_response(background_tasks, file)


@app.post("/convert/ppt-to-pdf")
async def convert_ppt_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    return await office_to_pdf_response(background_tasks, file)


@app.post("/convert/compress-pdf")
async def convert_compress_pdf(
    background_tasks: BackgroundTasks, file: UploadFile = File(...), level: str = Form("recommended")
):
    pdf_settings = {"low": "/printer", "recommended": "/ebook", "extreme": "/screen"}.get(level, "/ebook")
    work_dir = tempfile.mkdtemp()
    input_path = os.path.join(work_dir, "input.pdf")
    output_path = os.path.join(work_dir, "output.pdf")

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        before_size = os.path.getsize(input_path)

        result = subprocess.run(
            [
                GHOSTSCRIPT_BIN,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={pdf_settings}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                input_path,
            ],
            capture_output=True,
            timeout=100,
        )
        if result.returncode != 0 or not os.path.exists(output_path):
            stderr = result.stderr.decode(errors="ignore") if result.stderr else "unknown error"
            raise RuntimeError(f"Ghostscript compression failed: {stderr}")
        after_size = os.path.getsize(output_path)

        background_tasks.add_task(shutil.rmtree, work_dir, True)
        return FileResponse(
            path=output_path,
            filename="Compressed.pdf",
            media_type="application/pdf",
            headers={"X-Before-Size": str(before_size), "X-After-Size": str(after_size)},
        )
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/convert/pdf-thumbnail")
async def convert_pdf_thumbnail(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    fd_pdf, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
    fd_png, temp_png_path = tempfile.mkstemp(suffix=".png")
    os.close(fd_pdf)
    os.close(fd_png)

    try:
        with open(temp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = fitz.open(temp_pdf_path)
        if len(doc) == 0:
            raise RuntimeError("PDF has no pages")
        doc[0].get_pixmap(dpi=72).save(temp_png_path)

        background_tasks.add_task(remove_files, [temp_pdf_path, temp_png_path])
        return FileResponse(path=temp_png_path, media_type="image/png")
    except Exception as e:
        remove_files([temp_pdf_path, temp_png_path])
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/convert/office-thumbnail")
async def convert_office_thumbnail(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename or "")[1] or ".tmp"
    work_dir = tempfile.mkdtemp()
    input_path = os.path.join(work_dir, f"input{ext}")

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        pdf_path = convert_with_soffice(input_path, work_dir)
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            raise RuntimeError("Converted document has no pages")
        png_path = os.path.join(work_dir, "thumb.png")
        doc[0].get_pixmap(dpi=72).save(png_path)

        background_tasks.add_task(shutil.rmtree, work_dir, True)
        return FileResponse(path=png_path, media_type="image/png")
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))
