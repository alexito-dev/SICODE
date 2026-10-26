"""
BPSearch Router - Endpoints para búsqueda en Boletín El Peruano
"""

from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import Optional, List
import os
import json
import uuid
import shutil

from .processor import BPSearchProcessor
from app.modulos.auth.dependencies import login_required
from fastapi import Depends

router = APIRouter(prefix="/bpsearch", tags=["bpsearch"], dependencies=[Depends(login_required)])

from app.core.config import settings, sicode_dir

# Textos repetidos
_BUSQUEDA = "busqueda"
_CONTENT_DISPOSITION = "Content-Disposition"
_ERROR = "error"
_FOLDERS = "folders"
_MESSAGE = "message"
_NO_HAY_B_SQUEDA_ACTIVA = "No hay búsqueda activa"
_NO_HAY_RESULTADOS_PARA_EXPOR = "No hay resultados para exportar"
_STATUS = "status"
_SUCCESS = "success"

templates_dir = settings.get_template_path("app/modulos/bpsearch/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

# Instancia global del procesador (para demo - en producción usar sesiones)
processor_instance: Optional[BPSearchProcessor] = None
search_generator = None


class SearchRequest(BaseModel):
    carpeta: str
    palabra: str
    incluir_subcarpetas: bool = False
    ignorar_historial: bool = False


class FolderRequest(BaseModel):
    carpeta: str


# Define the central directory for BPSearch documents
BPSEARCH_BASE_DIR = sicode_dir("DATA", "BPSEARCH", "DOCS")


@router.get("/")
async def bpsearch_home(request: Request):
    return templates.TemplateResponse("bpsearch/index.html", {"request": request})


@router.post("/set-folder")
async def set_folder(data: FolderRequest):
    """Establece la subcarpeta de búsqueda dentro del directorio base o el directorio base mismo"""
    global processor_instance

    if processor_instance is None:
        processor_instance = BPSearchProcessor()

    # Resolve the physical path securely
    target_path = os.path.abspath(os.path.join(BPSEARCH_BASE_DIR, data.carpeta.strip(r"\/")))

    # Security check: ensure target_path is within BPSEARCH_BASE_DIR
    if not target_path.startswith(os.path.abspath(BPSEARCH_BASE_DIR)):
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: "Acceso denegado a esta ruta"})

    if processor_instance.set_carpeta(target_path):
        # Contar PDFs disponibles
        pdfs = processor_instance.obtener_pdfs(incluir_subcarpetas=False)
        pdfs_sub = processor_instance.obtener_pdfs(incluir_subcarpetas=True)

        display_name = "Directorio Principal" if data.carpeta == "" else data.carpeta

        return {
            _STATUS: _SUCCESS,
            "carpeta": display_name,
            "physical_path": target_path,
            "pdfs_carpeta": len(pdfs),
            "pdfs_total": len(pdfs_sub),
        }
    return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: "Esa carpeta no existe en el servidor"})


@router.post("/upload-pdfs")
async def upload_pdfs(files: List[UploadFile] = File(...)):
    """Sube archivos PDF desde el cliente y los pone en una carpeta de sesión"""
    global processor_instance
    if processor_instance is None:
        processor_instance = BPSearchProcessor()

    # Create a unique session folder for this user's upload
    session_id = str(uuid.uuid4())
    session_dir = os.path.join(BPSEARCH_BASE_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    uploaded_count = 0
    for file in files:
        if file.filename.lower().endswith(".pdf"):
            # Only use the basename to prevent path traversal
            safe_filename = os.path.basename(file.filename)
            file_path = os.path.join(session_dir, safe_filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            uploaded_count += 1

    if uploaded_count == 0:
        return JSONResponse(
            status_code=400, content={_STATUS: _ERROR, _MESSAGE: "No se encontraron PDFs válidos en la selección"}
        )

    processor_instance.set_carpeta(session_dir)
    return {
        _STATUS: _SUCCESS,
        "session_path": session_dir,
        "pdfs_carpeta": uploaded_count,
        "pdfs_total": uploaded_count,
    }


@router.post("/search")
async def start_search(data: SearchRequest):
    """Inicia una búsqueda SSE (Server-Sent Events)"""
    global processor_instance, search_generator

    if processor_instance is None:
        processor_instance = BPSearchProcessor()

    # Resolve the physical path securely
    target_path = data.carpeta
    if not os.path.isabs(target_path):
        target_path = os.path.abspath(os.path.join(BPSEARCH_BASE_DIR, target_path))

    if not target_path.startswith(os.path.abspath(BPSEARCH_BASE_DIR)):
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: "Acceso denegado a esta ruta"})

    if not processor_instance.set_carpeta(target_path):
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: "Carpeta no válida en el servidor"})

    # Crear el generador de búsqueda
    search_generator = processor_instance.buscar_palabra(
        palabra=data.palabra, incluir_subcarpetas=data.incluir_subcarpetas, ignorar_historial=data.ignorar_historial
    )

    return {_STATUS: "started", _MESSAGE: "Búsqueda iniciada"}


@router.get("/search-stream")
def search_stream():
    """Stream de eventos SSE para la búsqueda"""

    def event_generator():
        if search_generator is None:
            yield f"data: {json.dumps({'type': 'error', 'message': 'No hay búsqueda activa'})}\n\n"
            return

        try:
            for evento in search_generator:
                yield f"data: {json.dumps(evento)}\n\n"
                # No necesitamos sleep aquí si es síncrono en threadpool
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",  # Nginx
        },
    )


@router.post("/pause")
async def pause_search():
    """Pausa la búsqueda"""

    if processor_instance:
        processor_instance.pausar()
        return {_STATUS: "paused"}
    return {_STATUS: _ERROR, _MESSAGE: _NO_HAY_B_SQUEDA_ACTIVA}


@router.post("/resume")
async def resume_search():
    """Reanuda la búsqueda"""

    if processor_instance:
        processor_instance.continuar()
        return {_STATUS: "resumed"}
    return {_STATUS: _ERROR, _MESSAGE: _NO_HAY_B_SQUEDA_ACTIVA}


@router.post("/cancel")
async def cancel_search():
    """Cancela la búsqueda"""

    if processor_instance:
        processor_instance.cancelar()
        return {_STATUS: "cancelled"}
    return {_STATUS: _ERROR, _MESSAGE: _NO_HAY_B_SQUEDA_ACTIVA}


@router.get("/status")
async def get_status():
    """Obtiene el estado actual de la búsqueda"""

    if processor_instance:
        return processor_instance.get_estado()
    return {"estado": "idle", "progreso": 0, "total": 0}


@router.get("/results")
async def get_results():
    """Obtiene los resultados de la búsqueda"""

    if processor_instance and processor_instance.resultados:
        return {
            _STATUS: _SUCCESS,
            "count": len(processor_instance.resultados),
            "results": processor_instance.resultados,
        }
    return {_STATUS: "empty", "count": 0, "results": []}


@router.get("/export/txt")
async def export_txt(palabra: str = _BUSQUEDA):
    """Exporta resultados a TXT"""

    if not processor_instance or not processor_instance.resultados:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: _NO_HAY_RESULTADOS_PARA_EXPOR})

    contenido = processor_instance.exportar_txt(palabra)

    return StreamingResponse(
        iter([contenido.encode("utf-8")]),
        media_type="text/plain",
        headers={_CONTENT_DISPOSITION: f"attachment; filename=resultados_{palabra}.txt"},
    )


@router.get("/export/excel")
async def export_excel(palabra: str = _BUSQUEDA):
    """Exporta resultados a Excel"""

    if not processor_instance or not processor_instance.resultados:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: _NO_HAY_RESULTADOS_PARA_EXPOR})

    contenido = processor_instance.exportar_excel()

    return StreamingResponse(
        iter([contenido]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={_CONTENT_DISPOSITION: f"attachment; filename=resultados_{palabra}.xlsx"},
    )


@router.get("/export/pdf")
async def export_pdf(palabra: str = _BUSQUEDA):
    """Exporta resultados a PDF"""

    if not processor_instance or not processor_instance.resultados:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _MESSAGE: _NO_HAY_RESULTADOS_PARA_EXPOR})

    contenido = processor_instance.exportar_pdf(palabra)

    if not contenido:
        return JSONResponse(
            status_code=500, content={_STATUS: _ERROR, _MESSAGE: "Error generando PDF (¿fpdf instalado?)"}
        )

    return StreamingResponse(
        iter([contenido]),
        media_type="application/pdf",
        headers={_CONTENT_DISPOSITION: f"attachment; filename=resultados_{palabra}.pdf"},
    )


@router.post("/clear-history")
async def clear_history(data: dict):
    """Borra el historial de una palabra"""

    if not processor_instance:
        return {_STATUS: _ERROR, _MESSAGE: "Primero seleccione una carpeta"}

    palabra = data.get("palabra", "")
    if processor_instance.borrar_historial(palabra):
        return {_STATUS: _SUCCESS, _MESSAGE: f"Historial de '{palabra}' eliminado"}
    return {_STATUS: "info", _MESSAGE: "No existe historial para esta palabra"}


@router.get("/list-folders")
async def list_folders(path: str = ""):
    """Lista subcarpetas disponibles dentro del directorio central de BPSearch"""
    try:
        # Resolve target path securely
        target_path = os.path.abspath(os.path.join(BPSEARCH_BASE_DIR, path.strip(r"\/")))

        # Security check
        if not target_path.startswith(os.path.abspath(BPSEARCH_BASE_DIR)):
            return {_FOLDERS: [], _ERROR: "Acceso denegado"}

        if not os.path.isdir(target_path):
            return {_FOLDERS: [], _ERROR: "Ruta no válida en el servidor"}

        folders = []
        try:
            for item in os.listdir(target_path):
                full_path = os.path.join(target_path, item)
                if os.path.isdir(full_path):
                    # Solo enviamos el nombre o ruta relativa para mantener seguridad
                    rel_path = os.path.relpath(full_path, BPSEARCH_BASE_DIR)
                    folders.append({"name": item, "path": rel_path})
        except PermissionError:
            return {_FOLDERS: [], _ERROR: "Sin permisos en el servidor"}

        folders.sort(key=lambda x: x["name"].lower())

        # Determine relative current path for the frontend
        current_rel = (
            os.path.relpath(target_path, BPSEARCH_BASE_DIR) if target_path != os.path.abspath(BPSEARCH_BASE_DIR) else ""
        )

        return {_FOLDERS: folders, "current": current_rel}

    except Exception as e:
        return {_FOLDERS: [], _ERROR: str(e)}
