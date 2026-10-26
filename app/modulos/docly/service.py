import io
import os
import tempfile
import threading
import zipfile
from typing import Dict, Any, List

import fitz  # PyMuPDF
from docx2pdf import convert as _docx2pdf_convert

# Word (COM) solo admite una automatización a la vez en la misma sesión de
# Windows: dos conversiones simultáneas chocan entre sí y Word lanza errores
# como "Word.Application.Documents". Este lock serializa las conversiones
# para que, si dos usuarios (o dos disparos automáticos) coinciden, la
# segunda simplemente espere a que termine la primera en vez de fallar.
_word_lock = threading.Lock()


def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    """
    Convierte un Word (.docx/.doc) a PDF usando Microsoft Word instalado en la
    máquina (vía docx2pdf / COM). El resultado se ve exactamente igual que en
    Word: mismo formato, tablas, imágenes, estilos y saltos de página.
    """
    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "documento.docx")
        out_path = os.path.join(tmp, "documento.pdf")
        with open(in_path, "wb") as f:
            f.write(docx_bytes)

        with _word_lock:
            _docx2pdf_convert(in_path, out_path)

        with open(out_path, "rb") as f:
            return f.read()


def get_document_metadata(content: bytes, filename: str) -> Dict[str, Any]:
    """
    Inspecciona un archivo subido (.pdf o .docx) y retorna sus metadatos y total de páginas.
    """
    size_kb = round(len(content) / 1024, 2)
    ext = filename.split(".")[-1].lower() if "." in filename else ""

    total_pages = 0
    is_valid = True
    error_msg = None

    try:
        if ext == "pdf":
            doc = fitz.open(stream=content, filetype="pdf")
            total_pages = len(doc)
            doc.close()
        elif ext in ["docx", "doc"]:
            converted_pdf_bytes = convert_docx_to_pdf_bytes(content)
            doc = fitz.open(stream=converted_pdf_bytes, filetype="pdf")
            total_pages = len(doc)
            doc.close()
        else:
            is_valid = False
            error_msg = "Formato no soportado. Sube un archivo PDF (.pdf) o Word (.docx)."
    except Exception as e:
        is_valid = False
        error_msg = f"No se pudo leer el archivo {filename}: {str(e)}"

    return {
        "filename": filename,
        "extension": ext,
        "size_kb": size_kb,
        "is_valid": is_valid,
        "total_pages": max(1, total_pages),
        "error": error_msg,
    }


def subdividir_documento(
    content: bytes,
    filename: str,
    num_paginas: int,
    num_doc: int,
    nombres: str,
    pag_inicial: int = 1,
) -> List[Dict[str, Any]]:
    """
    Subdivide un documento (PDF o Word) en partes de N páginas.
    Devuelve la lista de sub-documentos generados, cada uno con su propio
    contenido en bytes, listo para descargarse por separado.
    """
    ext = filename.split(".")[-1].lower() if "." in filename else "pdf"

    # Si es Word, lo convertimos primero a PDF (con Word real) para
    # subdividir por página conservando el formato original.
    if ext in ["docx", "doc"]:
        pdf_bytes = convert_docx_to_pdf_bytes(content)
    else:
        pdf_bytes = content

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)

    if num_paginas <= 0:
        num_paginas = 1

    files: List[Dict[str, Any]] = []
    current_page = max(1, pag_inicial)
    doc_count = 1

    while current_page <= total_pages:
        if num_doc > 0 and doc_count > num_doc:
            break

        pagina_final = min(current_page + num_paginas - 1, total_pages)

        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=current_page - 1, to_page=pagina_final - 1)

        pdf_out = io.BytesIO()
        new_doc.save(pdf_out)
        new_doc.close()

        out_bytes = pdf_out.getvalue()
        files.append(
            {
                "filename": f"{nombres}{doc_count}.pdf",
                "content": out_bytes,
                "size_kb": round(len(out_bytes) / 1024, 1),
                "page_start": current_page,
                "page_end": pagina_final,
                "pages": pagina_final - current_page + 1,
            }
        )

        doc_count += 1
        current_page = pagina_final + 1

    doc.close()
    return files


def build_zip(files: List[Dict[str, Any]]) -> bytes:
    """Empaqueta la lista de sub-documentos en un único .zip (para 'descargar todo')."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for f in files:
            zip_file.writestr(f["filename"], f["content"])
    return zip_buffer.getvalue()
