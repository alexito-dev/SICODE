import datetime
from fastapi import APIRouter, Request, UploadFile, File, HTTPException, Depends
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse, Response
from sqlalchemy.orm import Session

from app.modulos.doccheck.processor import DocProcessor
from app.modulos.doccheck.models import DocCheckFile
from app.modulos.auth.dependencies import login_required
from app.core.database import get_db

router = APIRouter(prefix="/doccheck", tags=["doccheck"], dependencies=[Depends(login_required)])

from app.core.config import settings

# Textos repetidos
_ERROR = "error"
_NO_FILE_LOADED = "No file loaded"
_STATUS = "status"
_SUCCESS = "success"

templates_dir = settings.get_template_path("app/modulos/doccheck/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

# Cache en RAM por sesión activa: { user_id: DocProcessor }
# Evita releer la DB en cada petición durante la misma sesión del servidor
_cache: dict[int, DocProcessor] = {}


def _load_processor(user_id: int, db: Session) -> DocProcessor | None:
    """
    Carga el procesador para un usuario:
    1. Primero busca en RAM (_cache).
    2. Si no está, lo carga desde la DB.
    3. Si no tiene archivo guardado, retorna None.
    """
    if user_id in _cache:
        return _cache[user_id]

    record = db.query(DocCheckFile).filter(DocCheckFile.user_id == user_id).first()
    if record:
        processor = DocProcessor(record.file_data)
        _cache[user_id] = processor
        return processor

    return None


def _persist_processor(user_id: int, filename: str, processor: DocProcessor, db: Session):
    """
    Guarda/actualiza los bytes del workbook en la DB.
    Se llama después de cada upload o de cada save_record.
    """
    file_bytes = processor.get_workbook_bytes()
    record = db.query(DocCheckFile).filter(DocCheckFile.user_id == user_id).first()

    if record:
        record.file_data = file_bytes
        record.filename = filename
        record.updated_at = datetime.datetime.utcnow()
    else:
        record = DocCheckFile(
            user_id=user_id,
            filename=filename,
            file_data=file_bytes,
        )
        db.add(record)

    db.commit()


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------


@router.get("/")
async def doccheck_home(request: Request, db: Session = Depends(get_db)):
    user = request.state.user
    processor = _load_processor(user.id, db)
    has_file = processor is not None
    count = processor.record_count() if has_file else 0

    # Obtener nombre del archivo guardado (para mostrarlo en UI)
    filename = None
    if has_file:
        rec = db.query(DocCheckFile).filter(DocCheckFile.user_id == user.id).first()
        filename = rec.filename if rec else None

    return templates.TemplateResponse(
        "doccheck/index.html",
        {
            "request": request,
            "has_file": has_file,
            "record_count": count,
            "filename": filename,
        },
    )


@router.post("/upload")
async def upload_excel(request: Request, file: UploadFile = File(...), db: Session = Depends(get_db)):
    user = request.state.user
    try:
        file_bytes = await file.read()
        processor = DocProcessor(file_bytes)

        # Guardar en RAM
        _cache[user.id] = processor

        # Persistir en DB
        _persist_processor(user.id, file.filename or "doccheck.xlsx", processor, db)

        return {_STATUS: _SUCCESS, "count": processor.record_count()}
    except Exception as e:
        return JSONResponse(status_code=500, content={_STATUS: _ERROR, _ERROR: str(e)})


@router.get("/record/{idx}")
async def get_record(idx: int, request: Request, db: Session = Depends(get_db)):
    processor = _load_processor(request.state.user.id, db)
    if not processor:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _ERROR: _NO_FILE_LOADED})

    rec = processor.get_record(idx)
    if not rec:
        return JSONResponse(status_code=404, content={_STATUS: _ERROR, _ERROR: "Index out of range"})

    return {_STATUS: _SUCCESS, "data": rec}


@router.get("/find/{value}")
async def find_record(value: str, request: Request, db: Session = Depends(get_db)):
    """Buscar registro por N° de Registro y devolver su índice"""
    processor = _load_processor(request.state.user.id, db)
    if not processor:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _ERROR: _NO_FILE_LOADED})

    for idx in range(processor.record_count()):
        rec = processor.get_record(idx)
        if rec and str(rec.get("N° Registro", "")) == str(value):
            return {_STATUS: _SUCCESS, "index": idx}

    return {_STATUS: _ERROR, _ERROR: "Record not found"}


@router.post("/save/{idx}")
async def save_record(idx: int, request: Request, db: Session = Depends(get_db)):
    user = request.state.user
    processor = _load_processor(user.id, db)
    if not processor:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _ERROR: _NO_FILE_LOADED})

    try:
        form = await request.form()
        data = {k: v for k, v in form.items()}
        changed = processor.save_record(idx, data)

        if changed:
            # Actualizar DB con el workbook modificado
            rec = db.query(DocCheckFile).filter(DocCheckFile.user_id == user.id).first()
            filename = rec.filename if rec else "doccheck.xlsx"
            _persist_processor(user.id, filename, processor, db)

        return {_STATUS: _SUCCESS}
    except Exception as e:
        return JSONResponse(status_code=500, content={_STATUS: _ERROR, _ERROR: str(e)})


@router.get("/last-change")
async def last_change(request: Request, db: Session = Depends(get_db)):
    """Devuelve el último cambio registrado en el Historial del archivo del usuario."""
    processor = _load_processor(request.state.user.id, db)
    if not processor:
        return JSONResponse(status_code=400, content={_STATUS: _ERROR, _ERROR: _NO_FILE_LOADED})
    change = processor.get_last_change()
    if not change:
        return {_STATUS: "empty"}
    return {_STATUS: _SUCCESS, "data": change}


@router.get("/export/{format_type}")
async def export_data(format_type: str, request: Request, caja: str = "", db: Session = Depends(get_db)):
    processor = _load_processor(request.state.user.id, db)
    if not processor:
        raise HTTPException(status_code=400, detail=_NO_FILE_LOADED)

    try:
        if format_type == "excel":
            file_bytes = processor.export_excel_bytes(caja if caja else None)
            return Response(
                content=file_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": 'attachment; filename="Export.xlsx"'},
            )
        if format_type == "txt":
            file_bytes = processor.export_txt_bytes(caja if caja else None)
            return Response(
                content=file_bytes,
                media_type="text/plain",
                headers={"Content-Disposition": 'attachment; filename="Export.txt"'},
            )
        raise HTTPException(status_code=400, detail="Invalid format")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e
