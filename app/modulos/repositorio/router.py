"""
Módulo Repositorio - Router y endpoints para el almacenamiento centralizado
y distribución de archivos (ZIP, RAR, PDF, Excel, etc.) accesibles para todos los usuarios.
"""

import os
import re
import uuid
import urllib.parse
from datetime import datetime
from typing import List, Optional

from fastapi import APIRouter, Request, UploadFile, File, Form, Depends, HTTPException, Query
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import or_, desc, asc

from app.core.config import settings, sicode_dir
from app.core.database import SessionLocal
from app.modulos.auth.dependencies import login_required
from app.modulos.repositorio.models import RepoArchivo

# Textos repetidos
_STATUS = "status"
_SUCCESS = "success"

router = APIRouter(prefix="/repositorio", tags=["repositorio"], dependencies=[Depends(login_required)])

# Configuración de templates
templates_dir = settings.get_template_path("app/modulos/repositorio/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

# Directorio persistente de almacenamiento de archivos
REPO_DIR = sicode_dir("REPOSITORIO")


def format_file_size(size_in_bytes: int) -> str:
    """Convierte bytes a un formato legible por humanos (B, KB, MB, GB)."""
    if not size_in_bytes or size_in_bytes < 0:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.1f} {unit}" if unit != "B" else f"{int(size_in_bytes)} B"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.1f} PB"


def detect_file_category(extension: str) -> str:
    """Clasifica automáticamente el archivo según su extensión."""
    ext = (extension or "").lower().strip(".")

    categories = {
        "Comprimidos": ["zip", "rar", "7z", "tar", "gz", "bz2", "xz", "iso", "tgz", "z", "cab"],
        "Documentos": ["pdf", "doc", "docx", "odt", "rtf", "txt", "md", "epub"],
        "Hojas de Cálculo": ["xls", "xlsx", "xlsm", "csv", "tsv", "ods"],
        "Presentaciones": ["ppt", "pptx", "pps", "ppsx", "odp"],
        "Imágenes": ["png", "jpg", "jpeg", "gif", "svg", "webp", "bmp", "ico", "tiff"],
        "Multimedia": ["mp4", "avi", "mkv", "mov", "wmv", "mp3", "wav", "ogg", "flac", "m4a"],
        "Software / Scripts": ["exe", "msi", "bat", "cmd", "ps1", "sh", "py", "js", "sql", "json", "xml"],
    }

    for cat_name, ext_list in categories.items():
        if ext in ext_list:
            return cat_name
    return "Otros"


def sanitize_filename(filename: str) -> str:
    """Limpia caracteres potencialmente peligrosos de un nombre de archivo."""
    clean = re.sub(r'[\\/*?:"<>|]', "_", filename)
    clean = clean.strip()
    return clean or "archivo_sin_nombre"


@router.get("/")
async def repositorio_home(request: Request):
    """
    Renderiza la interfaz principal del Repositorio Compartido.
    """
    return templates.TemplateResponse(
        "repositorio/index.html", {"request": request, "title": "Repositorio Institucional de Archivos"}
    )


@router.get("/api/archivos")
async def listar_archivos(
    request: Request,
    q: Optional[str] = Query(None, description="Término de búsqueda"),
    categoria: Optional[str] = Query(None, description="Categoría a filtrar"),
    sort: Optional[str] = Query("recent", description="Criterio de ordenamiento"),
):
    """
    Retorna la lista de archivos disponibles en el repositorio con filtros y ordenación.
    """
    db = SessionLocal()
    try:
        query = db.query(RepoArchivo).filter(RepoArchivo.is_active)

        # Filtro de búsqueda por texto
        if q and q.strip():
            term = f"%{q.strip()}%"
            query = query.filter(
                or_(
                    RepoArchivo.filename_original.ilike(term),
                    RepoArchivo.description.ilike(term),
                    RepoArchivo.uploader_name.ilike(term),
                    RepoArchivo.extension.ilike(term),
                )
            )

        # Filtro por categoría
        if categoria and categoria != "Todos":
            query = query.filter(RepoArchivo.category == categoria)

        # Criterio de ordenación
        if sort == "oldest":
            query = query.order_by(asc(RepoArchivo.created_at))
        elif sort == "size_desc":
            query = query.order_by(desc(RepoArchivo.file_size))
        elif sort == "size_asc":
            query = query.order_by(asc(RepoArchivo.file_size))
        elif sort == "downloads":
            query = query.order_by(desc(RepoArchivo.downloads_count))
        elif sort == "name_asc":
            query = query.order_by(asc(RepoArchivo.filename_original))
        else:  # "recent" por defecto
            query = query.order_by(desc(RepoArchivo.created_at))

        archivos = query.all()

        # Determinar si el usuario actual es admin/supervisor para permisos de borrado
        current_user = getattr(request.state, "user", None)
        is_admin = False
        current_username = ""
        if current_user:
            is_admin = current_user.role in ["Administrador", "Supervisor"]
            current_username = current_user.username or ""

        results = []
        for a in archivos:
            data = a.to_dict()
            # Puede eliminar si es admin/supervisor o si es el dueño
            data["can_delete"] = is_admin or (current_username and a.uploader_username == current_username)
            results.append(data)

        return {_STATUS: _SUCCESS, "count": len(results), "data": results}
    finally:
        db.close()


@router.post("/api/upload")
async def subir_archivos(
    request: Request,
    files: List[UploadFile] = File(...),
    categoria: Optional[str] = Form("Auto"),
    descripcion: Optional[str] = Form(""),
):
    """
    Guarda uno o múltiples archivos subidos en el almacenamiento persistente
    y registra sus metadatos en la base de datos de SICODE.
    """
    if not files:
        raise HTTPException(status_code=400, detail="No se seleccionó ningún archivo para subir.")

    # Información del usuario que realiza la carga
    user = getattr(request.state, "user", None)
    uploader_name = user.full_name if user else "Invitado"
    uploader_username = user.username if user else "invitado"
    uploader_id = user.id if user else None

    db = SessionLocal()
    uploaded_records = []

    try:
        for file_item in files:
            if not file_item.filename:
                continue

            orig_name = sanitize_filename(file_item.filename)
            _, ext = os.path.splitext(orig_name)
            ext_clean = ext.lower().replace(".", "")

            # Determinar categoría
            if not categoria or categoria == "Auto":
                final_cat = detect_file_category(ext_clean)
            else:
                final_cat = categoria

            # Generar nombre único para el archivo en disco (evita colisiones)
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            unique_prefix = uuid.uuid4().hex[:8]
            storage_name = f"{timestamp_str}_{unique_prefix}_{orig_name}"
            dest_path = os.path.join(REPO_DIR, storage_name)

            # Escribir el archivo por bloques para soportar archivos pesados (ZIP, RAR, etc.) sin consumir toda la RAM
            total_bytes = 0
            with open(dest_path, "wb") as buffer:
                while chunk := await file_item.read(1024 * 1024):  # Bloques de 1 MB
                    buffer.write(chunk)
                    total_bytes += len(chunk)

            formatted_size = format_file_size(total_bytes)

            # Crear registro en la BD
            nuevo_archivo = RepoArchivo(
                filename_original=orig_name,
                filename_storage=storage_name,
                file_path=dest_path,
                file_size=total_bytes,
                file_size_formatted=formatted_size,
                extension=ext_clean,
                category=final_cat,
                description=descripcion.strip() if descripcion else None,
                uploader_id=uploader_id,
                uploader_name=uploader_name,
                uploader_username=uploader_username,
                downloads_count=0,
                created_at=datetime.now(),
                is_active=True,
            )

            db.add(nuevo_archivo)
            db.commit()
            db.refresh(nuevo_archivo)

            uploaded_records.append(nuevo_archivo.to_dict())

        return {
            _STATUS: _SUCCESS,
            "message": f"Se subieron {len(uploaded_records)} archivo(s) correctamente.",
            "data": uploaded_records,
        }

    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Error al guardar los archivos: {str(e)}") from e
    finally:
        db.close()


@router.get("/download/{file_id}")
async def descargar_archivo(file_id: int):
    """
    Descarga el archivo solicitado e incrementa el contador de descargas.
    """
    db = SessionLocal()
    try:
        archivo = db.query(RepoArchivo).filter(RepoArchivo.id == file_id, RepoArchivo.is_active).first()

        if not archivo:
            raise HTTPException(status_code=404, detail="El archivo solicitado no existe o fue retirado.")

        if not os.path.exists(archivo.file_path):
            raise HTTPException(status_code=404, detail="El archivo físico no fue encontrado en el disco.")

        # Incrementar contador de descargas
        archivo.downloads_count = (archivo.downloads_count or 0) + 1
        db.commit()

        # Codificar nombre seguro para Content-Disposition
        safe_filename = archivo.filename_original
        encoded_filename = urllib.parse.quote(safe_filename)

        headers = {"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}

        return FileResponse(
            path=archivo.file_path, filename=safe_filename, media_type="application/octet-stream", headers=headers
        )
    finally:
        db.close()


@router.get("/preview/{file_id}")
async def previsualizar_archivo(file_id: int):
    """
    Permite visualizar directamente en el navegador archivos compatibles (PDF, imágenes, texto).
    """
    db = SessionLocal()
    try:
        archivo = db.query(RepoArchivo).filter(RepoArchivo.id == file_id, RepoArchivo.is_active).first()

        if not archivo or not os.path.exists(archivo.file_path):
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")

        ext = (archivo.extension or "").lower()
        media_types = {
            "pdf": "application/pdf",
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "svg": "image/svg+xml",
            "webp": "image/webp",
            "txt": "text/plain; charset=utf-8",
            "csv": "text/plain; charset=utf-8",
        }

        content_type = media_types.get(ext, "application/octet-stream")
        return FileResponse(path=archivo.file_path, media_type=content_type)
    finally:
        db.close()


@router.delete("/api/archivos/{file_id}")
async def eliminar_archivo(file_id: int, request: Request):
    """
    Elimina un archivo del repositorio físico y de la base de datos.
    Solo permitido para el usuario que lo subió o usuarios Administrador/Supervisor.
    """
    user = getattr(request.state, "user", None)
    if not user:
        raise HTTPException(status_code=401, detail="Sesión no iniciada.")

    db = SessionLocal()
    try:
        archivo = db.query(RepoArchivo).filter(RepoArchivo.id == file_id).first()
        if not archivo:
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")

        # Verificar permisos
        is_admin = user.role in ["Administrador", "Supervisor"]
        is_owner = archivo.uploader_username == user.username

        if not (is_admin or is_owner):
            raise HTTPException(status_code=403, detail="No tienes permisos para eliminar este archivo.")

        # Eliminar archivo físico de disco de forma segura
        try:
            if os.path.exists(archivo.file_path):
                os.remove(archivo.file_path)
        except Exception:
            # Continuar incluso si el archivo no estaba físicamente
            pass

        # Eliminar de la base de datos
        db.delete(archivo)
        db.commit()

        return {_STATUS: _SUCCESS, "message": f"Archivo '{archivo.filename_original}' eliminado correctamente."}
    finally:
        db.close()


@router.get("/api/stats")
async def obtener_estadisticas():
    """
    Retorna métricas globales del repositorio: total archivos, espacio ocupado,
    descargas totales y distribución por categorías.
    """
    db = SessionLocal()
    try:
        archivos = db.query(RepoArchivo).filter(RepoArchivo.is_active).all()

        total_count = len(archivos)
        total_bytes = sum(a.file_size for a in archivos)
        total_downloads = sum(a.downloads_count or 0 for a in archivos)

        # Conteo por categoría
        categories_count = {}
        for a in archivos:
            cat = a.category or "Otros"
            categories_count[cat] = categories_count.get(cat, 0) + 1

        return {
            _STATUS: _SUCCESS,
            "total_files": total_count,
            "total_bytes": total_bytes,
            "total_size_formatted": format_file_size(total_bytes),
            "total_downloads": total_downloads,
            "categories": categories_count,
        }
    finally:
        db.close()
