import base64

from fastapi import APIRouter, Request, Depends, UploadFile, File, Form
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse
from fastapi.concurrency import run_in_threadpool
from app.modulos.auth.dependencies import login_required
from app.core.config import settings
from app.modulos.docly.service import (
    get_document_metadata,
    subdividir_documento,
    build_zip,
)

router = APIRouter(prefix="/docly", tags=["docly"], dependencies=[Depends(login_required)])

templates_dir = settings.get_template_path("app/modulos/docly/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get("/")
async def docly_home(request: Request):
    """Renderiza la interfaz principal del módulo Docly."""
    return templates.TemplateResponse(
        "docly/index.html", {"request": request, "titulo": "Docly - Subdivisor de PDF y Word"}
    )


@router.post("/api/doc-info")
async def get_doc_info(file: UploadFile = File(...)):
    """Obtiene metadatos y número de páginas del documento subido (PDF o Word)."""
    if not file.filename:
        return JSONResponse(status_code=400, content={"success": False, "error": "No se recibió un archivo válido."})

    ext = file.filename.split(".")[-1].lower() if "." in file.filename else ""
    if ext not in ["pdf", "docx", "doc"]:
        return JSONResponse(
            status_code=400,
            content={
                "success": False,
                "error": "Formato no permitido. Solo se admiten archivos PDF (.pdf) o Word (.docx, .doc).",
            },
        )

    try:
        content = await file.read()
        # La conversión con Word (COM) es bloqueante: se corre en threadpool
        # para no congelar el servidor mientras procesa.
        meta = await run_in_threadpool(get_document_metadata, content, file.filename)
        if not meta["is_valid"]:
            return JSONResponse(status_code=400, content={"success": False, "error": meta["error"]})

        return {"success": True, **meta}
    except Exception as e:
        return JSONResponse(
            status_code=500, content={"success": False, "error": f"Error al leer el documento: {str(e)}"}
        )


@router.post("/api/subdividir")
async def subdividir_documento_endpoint(
    file: UploadFile = File(...),
    num_paginas: int = Form(1),
    num_doc: int = Form(0),
    nombres: str = Form("Documento_"),
    pag_inicial: int = Form(1),
):
    """
    Subdivide un documento en sub-archivos PDF. Devuelve cada sub-documento
    en base64 (para descargarlo individualmente) además de un .zip con todos.
    """
    if not file.filename:
        return JSONResponse(status_code=400, content={"success": False, "error": "Se requiere un archivo válido."})

    ext = file.filename.split(".")[-1].lower() if "." in file.filename else ""
    if ext not in ["pdf", "docx", "doc"]:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Formato no permitido. Sube un PDF (.pdf) o Word (.docx, .doc)."},
        )

    if num_paginas <= 0:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "El número de páginas por documento debe ser al menos 1."},
        )

    try:
        content = await file.read()
        files = await run_in_threadpool(
            subdividir_documento,
            content=content,
            filename=file.filename,
            num_paginas=num_paginas,
            num_doc=num_doc,
            nombres=nombres,
            pag_inicial=pag_inicial,
        )

        if not files:
            return JSONResponse(
                status_code=400,
                content={"success": False, "error": "No se generó ningún documento con estos parámetros."},
            )

        safe_prefix = "".join(c for c in nombres if c.isalnum() or c in ("_", "-")).rstrip("_")
        zip_filename = f"Docly_{safe_prefix or 'Divididos'}.zip"
        zip_bytes = build_zip(files)

        return {
            "success": True,
            "zip_filename": zip_filename,
            "zip_base64": base64.b64encode(zip_bytes).decode(),
            "files": [
                {
                    "filename": f["filename"],
                    "size_kb": f["size_kb"],
                    "page_start": f["page_start"],
                    "page_end": f["page_end"],
                    "pages": f["pages"],
                    "base64": base64.b64encode(f["content"]).decode(),
                }
                for f in files
            ],
        }
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": f"Error al procesar y subdividir el documento: {str(e)}"},
        )
