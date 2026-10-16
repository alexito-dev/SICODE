import os
import math
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import io, base64, json
from app.modulos.auth.dependencies import login_required

router = APIRouter(
    prefix="/datamask",
    tags=["datamask"]
)

from app.core.config import settings
templates_dir = settings.get_template_path("app/modulos/datamask/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get("/")
async def datamask_home(request: Request):
    return templates.TemplateResponse("datamask/index.html", {"request": request})


# ─── Process (annotate/mask/images/text) ───────────────────────────────────────


def safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except (ValueError, TypeError):
        return default


@router.post("/process")
async def process_pdf(
    file: UploadFile = File(...),
    annotations: str = Form(...),
    pageRotations: str = Form("{}")
):
    try:
        try:
            data = json.loads(annotations or "{}")
        except Exception:
            data = {}

        try:
            user_page_rotations = json.loads(pageRotations or "{}")
        except Exception:
            user_page_rotations = {}

        pdf_bytes = await file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        if user_page_rotations and isinstance(user_page_rotations, dict):
            for p_idx_s, add_rot in user_page_rotations.items():
                try:
                    p_idx = int(p_idx_s)
                    if 0 <= p_idx < len(doc):
                        doc[p_idx].set_rotation((doc[p_idx].rotation + int(add_rot)) % 360)
                except Exception:
                    pass

        if isinstance(data, dict):
            for page_idx_str, anns in data.items():
                try:
                    page_idx = int(page_idx_str)
                except ValueError:
                    continue

                if page_idx >= len(doc):
                    continue

                if not anns or not isinstance(anns, list):
                    continue

                page = doc[page_idx]
                pw, ph = page.rect.width, page.rect.height
                page_rot = page.rotation % 360
                derot = page.derotation_matrix
                shape = page.new_shape()
                has_shape_drawings = False
                has_redactions = False

                for ann in anns:
                    if not isinstance(ann, dict):
                        continue
                    try:
                        kind = ann.get("type")
                        hex_c = ann.get("color", "#000000")
                        if isinstance(hex_c, str) and hex_c.startswith("#") and len(hex_c) == 7:
                            try:
                                r = int(hex_c[1:3], 16) / 255.0
                                g = int(hex_c[3:5], 16) / 255.0
                                b = int(hex_c[5:7], 16) / 255.0
                            except ValueError:
                                r, g, b = 0.0, 0.0, 0.0
                        else:
                            r, g, b = 0.0, 0.0, 0.0
                        col = (r, g, b)

                        if kind == "highlight":
                            r_list = ann.get("rect") or [0.1, 0.1, 0.2, 0.1]
                            rx = safe_float(r_list[0] if len(r_list) > 0 else 0.1)
                            ry = safe_float(r_list[1] if len(r_list) > 1 else 0.1)
                            rw = safe_float(r_list[2] if len(r_list) > 2 else 0.2)
                            rh = safe_float(r_list[3] if len(r_list) > 3 else 0.1)
                            r_vis = fitz.Rect(page.rect.x0 + rx * pw, page.rect.y0 + ry * ph, page.rect.x0 + (rx + rw) * pw, page.rect.y0 + (ry + rh) * ph)
                            rect = (r_vis * derot).normalize()
                            shape.draw_rect(rect)
                            shape.finish(fill=col, fill_opacity=0.35)
                            has_shape_drawings = True

                        elif kind == "redact":
                            r_list = ann.get("rect") or [0.1, 0.1, 0.2, 0.1]
                            rx = safe_float(r_list[0] if len(r_list) > 0 else 0.1)
                            ry = safe_float(r_list[1] if len(r_list) > 1 else 0.1)
                            rw = safe_float(r_list[2] if len(r_list) > 2 else 0.2)
                            rh = safe_float(r_list[3] if len(r_list) > 3 else 0.1)
                            r_vis = fitz.Rect(page.rect.x0 + rx * pw, page.rect.y0 + ry * ph, page.rect.x0 + (rx + rw) * pw, page.rect.y0 + (ry + rh) * ph)
                            rect = (r_vis * derot).normalize()
                            page.add_redact_annot(rect, fill=col)
                            has_redactions = True

                        elif kind == "strike":
                            p1_dict = ann.get("p1") if isinstance(ann.get("p1"), dict) else {"x": 0.1, "y": 0.1}
                            p2_dict = ann.get("p2") if isinstance(ann.get("p2"), dict) else {"x": 0.2, "y": 0.2}
                            pt1 = fitz.Point(page.rect.x0 + safe_float(p1_dict.get("x"), 0.1) * pw, page.rect.y0 + safe_float(p1_dict.get("y"), 0.1) * ph) * derot
                            pt2 = fitz.Point(page.rect.x0 + safe_float(p2_dict.get("x"), 0.2) * pw, page.rect.y0 + safe_float(p2_dict.get("y"), 0.2) * ph) * derot
                            shape.draw_line(pt1, pt2)
                            shape.finish(color=col, width=safe_float(ann.get("size"), 2))
                            has_shape_drawings = True

                        elif kind == "ink":
                            pts = ann.get("points") if isinstance(ann.get("points"), list) else []
                            if len(pts) > 1:
                                valid_pts = []
                                for p in pts:
                                    if isinstance(p, dict):
                                        valid_pts.append(fitz.Point(page.rect.x0 + safe_float(p.get("x"), 0.1) * pw, page.rect.y0 + safe_float(p.get("y"), 0.1) * ph) * derot)
                                if len(valid_pts) > 1:
                                    shape.draw_polyline(valid_pts)
                                    shape.finish(color=col, width=safe_float(ann.get("size"), 3))
                                    has_shape_drawings = True

                        elif kind == "text":
                            tx = page.rect.x0 + safe_float(ann.get("x"), 0.1) * pw
                            ty = page.rect.y0 + safe_float(ann.get("y"), 0.1) * ph
                            tw = safe_float(ann.get("w"), 0.35) * pw
                            th = safe_float(ann.get("h"), 0.08) * ph
                            f_size = safe_float(ann.get("fontSize"), 13)
                            ann_rot = safe_float(ann.get("rotation"), 0)
                            total_rot = int((page_rot + ann_rot) % 360)

                            text_str = str(ann.get("text") or "").strip()
                            if text_str:
                                box_w = max(tw, 40.0)
                                box_h = max(th, 15.0)
                                r_vis = fitz.Rect(tx, ty, tx + box_w, ty + box_h)
                                t_rect = (r_vis * derot).normalize()
                                rc = page.insert_textbox(t_rect, text_str, fontsize=f_size, color=col, rotate=total_rot)
                                if rc < 0:
                                    r_vis = fitz.Rect(tx, ty, tx + box_w + 200.0, ty + box_h + 120.0)
                                    t_rect = (r_vis * derot).normalize()
                                    page.insert_textbox(t_rect, text_str, fontsize=max(8.0, f_size - 2), color=col, rotate=total_rot)

                        elif kind == "image":
                            img_data_str = str(ann.get("image") or "")
                            ann_rot = safe_float(ann.get("rotation"), 0)
                            total_rot = int((page_rot + ann_rot) % 360)

                            if img_data_str:
                                if "," in img_data_str:
                                    img_data_str = img_data_str.split(",", 1)[1]
                                try:
                                    raw_img = base64.b64decode(img_data_str)
                                    img_obj = Image.open(io.BytesIO(raw_img))
                                    img_obj = ImageOps.exif_transpose(img_obj)
                                    if img_obj.mode != "RGBA":
                                        img_obj = img_obj.convert("RGBA")

                                    img_byte_arr = io.BytesIO()
                                    img_obj.save(img_byte_arr, format="PNG")
                                    png_bytes = img_byte_arr.getvalue()

                                    ix = page.rect.x0 + safe_float(ann.get("x"), 0.1) * pw
                                    iy = page.rect.y0 + safe_float(ann.get("y"), 0.1) * ph
                                    iw = max(10.0, safe_float(ann.get("w"), 0.3) * pw)
                                    ih = max(10.0, safe_float(ann.get("h"), 0.15) * ph)

                                    if total_rot in [90, 270]:
                                        cx, cy = ix + iw / 2.0, iy + ih / 2.0
                                        r_vis = fitz.Rect(cx - ih / 2.0, cy - iw / 2.0, cx + ih / 2.0, cy + iw / 2.0)
                                    else:
                                        r_vis = fitz.Rect(ix, iy, ix + iw, iy + ih)

                                    img_rect = (r_vis * derot).normalize()
                                    page.insert_image(img_rect, stream=png_bytes, rotate=total_rot, keep_proportion=False, overlay=True)
                                except Exception as img_err:
                                    print(f"Error inserting image on page {page_idx}: {img_err}")
                    except Exception as ann_item_err:
                        print(f"Error processing ann item on page {page_idx}: {ann_item_err}")

                if has_shape_drawings:
                    try:
                        shape.commit()
                    except Exception as shape_err:
                        print(f"Shape commit error on page {page_idx}: {shape_err}")

                if has_redactions:
                    try:
                        page.apply_redactions()
                    except Exception as red_err:
                        print(f"Redaction error on page {page_idx}: {red_err}")

        buf = io.BytesIO()
        doc.save(buf)
        doc.close()
        buf.seek(0)
        pdf_bytes = buf.getvalue()
        pdf_b64 = base64.b64encode(pdf_bytes).decode()

        # Direct Desktop save fallback
        try:
            filename = file.filename or "documento.pdf"
            base_name = os.path.splitext(filename)[0] + "_editado.pdf"
            if not base_name.endswith(".pdf"):
                base_name += ".pdf"

            escritorio_onedrive = os.path.join("C:", os.sep, "Users", "alets", "OneDrive", "Escritorio", base_name)
            escritorio_local = os.path.join("C:", os.sep, "Users", "alets", "Desktop", base_name)

            for p in [escritorio_onedrive, escritorio_local]:
                try:
                    os.makedirs(os.path.dirname(p), exist_ok=True)
                    with open(p, "wb") as f_out:
                        f_out.write(pdf_bytes)
                    print(f"=== SUCESSFULLY SAVED TO DESKTOP: {p} ===")
                except Exception as ex:
                    print(f"Desktop path save warning: {ex}")
        except Exception as dt_err:
            print(f"Desktop save warning: {dt_err}")

        return JSONResponse(content={
            "status": "success",
            "file": pdf_b64
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return JSONResponse(status_code=200, content={"status": "error", "error": str(e)})

@router.post("/save_to_desktop")
async def save_to_desktop(
    file: UploadFile = File(...),
    annotations: str = Form(...),
    pageRotations: str = Form("{}")
):
    try:
        res = await process_pdf(file=file, annotations=annotations, pageRotations=pageRotations)
        body = json.loads(res.body.decode())
        if body.get("status") == "success" and body.get("file"):
            pdf_b64 = body["file"]
            pdf_bytes = base64.b64decode(pdf_b64)

            filename = file.filename or "documento.pdf"
            base_name = os.path.splitext(filename)[0] + "_editado.pdf"
            if not base_name.endswith(".pdf"):
                base_name += ".pdf"

            escritorio_onedrive = os.path.join("C:", os.sep, "Users", "alets", "OneDrive", "Escritorio", base_name)
            escritorio_local = os.path.join("C:", os.sep, "Users", "alets", "Desktop", base_name)

            paths_to_save = [
                escritorio_onedrive,
                escritorio_local,
                os.path.join(os.path.expanduser("~"), "OneDrive", "Escritorio", base_name),
                os.path.join(os.path.expanduser("~"), "Desktop", base_name)
            ]

            saved_paths = []
            for p in paths_to_save:
                try:
                    os.makedirs(os.path.dirname(p), exist_ok=True)
                    with open(p, "wb") as f_out:
                        f_out.write(pdf_bytes)
                    saved_paths.append(p)
                except Exception:
                    pass

            return JSONResponse(content={
                "status": "success",
                "message": f"Guardado en Escritorio como {base_name}",
                "paths": saved_paths,
                "file": pdf_b64
            })
        else:
            return res
    except Exception as e:
        return JSONResponse(status_code=200, content={"status": "error", "error": str(e)})

@router.post("/convert")
async def convert_pdf_to_word(
    file: UploadFile = File(...),
    ocr: str = Form("1"),
    layout: str = Form("1"),
):
    """
    Convierte PDF a DOCX usando pdf2docx (texto nativo) y
    easyocr con fallback a pdf2docx para páginas escaneadas.
    """
    try:
        pdf_bytes = await file.read()
        use_ocr = ocr == "1"

        # ── Detectar si el PDF tiene texto nativo ──
        src_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_chars = sum(len(src_doc[i].get_text()) for i in range(min(3, len(src_doc))))
        is_scanned = total_chars < 100  # muy poco texto → escaneado

        docx_bytes = None
        method = "native"

        if is_scanned and use_ocr:
            try:
                docx_bytes = _ocr_to_docx(src_doc)
                method = "ocr"
            except Exception as ocr_err:
                print(f"OCR no disponible o falló ({ocr_err}), usando pdf2docx como fallback...")
                docx_bytes = _pdf2docx_convert(pdf_bytes)
                method = "native_fallback"
        else:
            docx_bytes = _pdf2docx_convert(pdf_bytes)

        src_doc.close()
        return JSONResponse(content={
            "status": "success",
            "file": base64.b64encode(docx_bytes).decode(),
            "method": method
        })

    except ImportError as e:
        return JSONResponse(status_code=500, content={
            "error": f"Librería requerida: {e}. Ejecuta: pip install pdf2docx"
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return JSONResponse(status_code=200, content={"status": "error", "error": str(e)})


def _pdf2docx_convert(pdf_bytes: bytes) -> bytes:
    """Usa pdf2docx para convertir PDF a DOCX de manera segura en Windows."""
    from pdf2docx import Converter
    import tempfile, os

    pdf_fd, pdf_path = tempfile.mkstemp(suffix=".pdf")
    os.close(pdf_fd)
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)

    out_path = pdf_path.replace(".pdf", ".docx")
    try:
        cv = Converter(pdf_path)
        cv.convert(out_path, start=0, end=None)
        cv.close()
        with open(out_path, "rb") as f:
            return f.read()
    finally:
        for p in (pdf_path, out_path):
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except Exception:
                pass


def _ocr_to_docx(src_doc) -> bytes:
    """
    Para PDFs escaneados: usa EasyOCR si está disponible para extraer texto
    y genera un documento DOCX.
    """
    import easyocr
    from docx import Document
    from docx.shared import Pt
    import numpy as np
    import cv2

    reader = easyocr.Reader(['es', 'en'], gpu=False)
    out_doc = Document()
    out_doc.add_heading('Documento Convertido (OCR)', 0)

    for page_num in range(len(src_doc)):
        page = src_doc[page_num]

        # Renderizar a imagen alta res para mejor OCR
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        img_arr = np.frombuffer(img_bytes, dtype=np.uint8)

        img = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
        results = reader.readtext(img, detail=1, paragraph=False)

        if page_num > 0:
            out_doc.add_page_break()

        out_doc.add_heading(f'Página {page_num + 1}', level=2)

        if results:
            results_sorted = sorted(results, key=lambda x: x[0][0][1])  # sort by y
            current_line_y = results_sorted[0][0][0][1]
            line_texts = []

            for bbox, text, conf in results_sorted:
                y = bbox[0][1]
                if abs(y - current_line_y) > 20:
                    if line_texts:
                        p = out_doc.add_paragraph(' '.join(line_texts))
                        if p.runs:
                            p.runs[0].font.size = Pt(11)
                    line_texts = [text]
                    current_line_y = y
                else:
                    line_texts.append(text)

            if line_texts:
                out_doc.add_paragraph(' '.join(line_texts))

    buf = io.BytesIO()
    out_doc.save(buf)
    buf.seek(0)
    return buf.getvalue()