from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pdf2docx import Converter
import pdfplumber
from openpyxl import Workbook
from pptx import Presentation
import fitz # PyMuPDF
import tempfile
import os
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


def convert_with_soffice(input_path: str, output_dir: str) -> str:
    """Converts a document to PDF via headless LibreOffice. Used for word/excel/ppt -> pdf."""
    result = subprocess.run(
        [SOFFICE_BIN, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", output_dir, input_path],
        capture_output=True,
        timeout=100,
    )
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.pdf")
    if result.returncode != 0 or not os.path.exists(output_path):
        stderr = result.stderr.decode(errors="ignore") if result.stderr else "unknown error"
        raise RuntimeError(f"LibreOffice conversion failed: {stderr}")
    return output_path


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

        def convert_with_correct_spacing(pdf_path, docx_path):
            out_doc = docx.Document()
            pdf = fitz.open(pdf_path)
            for page in pdf:
                blocks = page.get_text("dict")["blocks"]
                all_lines = []
                for b in blocks:
                    if b.get("type") == 0:
                        for l in b.get("lines", []):
                            spans = l.get("spans", [])
                            if not spans: continue
                            
                            run_params = []
                            last_x = None
                            first_span = spans[0]
                            font_size = first_span.get("size", 11)
                            
                            for s in spans:
                                text = s.get("text", "")
                                bbox = s.get("bbox", [0,0,0,0])
                                
                                if last_x is not None:
                                    gap = bbox[0] - last_x
                                    if gap > 2 and not text.startswith(" ") and not text.startswith(","):
                                        run_params.append({"text": " ", "size": font_size, "font": first_span.get("font", ""), "flags": 0})
                                
                                run_params.append({"text": text, "size": s.get("size", 11), "font": s.get("font", "").lower(), "flags": s.get("flags", 0)})
                                last_x = bbox[2]
                                
                            full_text = "".join(r["text"] for r in run_params).strip()
                            all_lines.append({
                                "bbox": l.get("bbox", [0,0,0,0]),
                                "runs": run_params,
                                "full_text": full_text,
                                "font_size": font_size
                            })
                            
                # Sort lines by approximate Y, then X
                all_lines.sort(key=lambda x: (round(x["bbox"][1] / 5), x["bbox"][0]))
                
                p = None
                last_y_bottom = -999
                last_font_size = -1
                
                for line in all_lines:
                    if not line["full_text"]: continue
                    
                    bbox = line["bbox"]
                    y_top = bbox[1]
                    y_bottom = bbox[3]
                    font_size = line["font_size"]
                    vertical_gap = y_top - last_y_bottom
                    
                    is_bullet = line["full_text"].startswith(('•', '◦', '-', '*'))
                    
                    start_new_para = True
                    if p is not None:
                        if is_bullet:
                            start_new_para = True
                        elif y_top < (last_y_bottom - font_size * 0.3):
                            # Column-based content (same row, right aligned)
                            start_new_para = True
                        elif vertical_gap < (font_size * 0.5) and abs(font_size - last_font_size) < 1.0:
                            start_new_para = False
                            
                    if start_new_para:
                        p = out_doc.add_paragraph()
                    else:
                        p.add_run(" ")
                        
                    import re
                    def clean_font_name(f_str):
                        name = f_str.split('+')[-1]
                        name = re.sub(r'-(Bold|Italic|Regular|Medium|SemiBold|Light).*', '', name, flags=re.IGNORECASE)
                        name = re.sub(r'(MT|PSMT|MS|PS)$', '', name, flags=re.IGNORECASE)
                        name = re.sub(r'([a-z])([A-Z])', r'\1 \2', name)
                        return name.strip()

                    for r in line["runs"]:
                        run = p.add_run(r["text"])
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
                            run.font.size = docx.shared.Pt(round(r["size"]))
                            
                    last_y_bottom = y_bottom
                    last_font_size = font_size
            out_doc.save(docx_path)
        
        # 1. Primary conversion using pdf2docx (Best for formatting)
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()
        
        # 2. Check and fallback (Fixes spacing issues on Canva PDFs with heuristics)
        used_method = "layout-mode"
        if has_spacing_issue(temp_docx_path):
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
        import pdfplumber
        
        wb = Workbook()
        ws = wb.active
        
        with pdfplumber.open(temp_pdf_path) as pdf:
            tables_found = False
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        tables_found = True
                        for row in table:
                            clean_row = [cell if cell is not None else "" for cell in row]
                            ws.append(clean_row)
                        ws.append([]) # empty row between tables
            
            # Fallback: if no tables are found (e.g. resumes), extract text line-by-line
            if not tables_found:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            ws.append([line.strip()])
                        ws.append([])
                        
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
        return {"error": str(e)}

@app.post("/convert/speech-to-text")
async def convert_speech_to_text(file: UploadFile = File(...)):
    # Save the uploaded file temporarily
    fd_audio, temp_audio_path = tempfile.mkstemp(suffix=os.path.splitext(file.filename)[1])
    os.close(fd_audio)

    try:
        with open(temp_audio_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Transcribe the audio
        segments, info = whisper_model.transcribe(temp_audio_path, beam_size=5)
        
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
