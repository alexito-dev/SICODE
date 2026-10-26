import io
import os
import json
import shutil
from datetime import datetime
from fastapi import APIRouter, Request, Depends, UploadFile, File
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse, Response
from app.modulos.auth.dependencies import login_required
from app.core.config import settings, DB_DIR
from app.modulos.odi.odi_service import analyze_excel
from app.modulos.odi.odi_export import generate_excel_export, generate_pdf_export
import contextlib

# Textos repetidos
_ACCESS_CONTROL_EXPOSE_HEADER = "Access-Control-Expose-Headers"
_COLABORADOR = "Colaborador"
_CONTENT_DISPOSITION = "Content-Disposition"
_DB_DIR = "DB_DIR"
_ERROR = "error"
_FILENAME = "filename"
_FILES = "files"
_FULL_NAME = "full_name"
_RUC_MAP = "ruc_map"
_SHEET = "sheet"
_SHEETS = "sheets"
_SIZE_KB = "size_kb"
_TOTAL_RUCS = "total_rucs"
_UPLOADED_AT = "uploaded_at"
_UPLOADED_BY = "uploaded_by"
_UPLOADER_ID = "uploader_id"
_UPLOADER_ROLE = "uploader_role"
_USERNAME = "username"
_USUARIO = "Usuario"
_UTF_8 = "utf-8"
_Y_M_D__H_M_S = "%Y%m%d_%H%M%S"

router = APIRouter(prefix="/odi", tags=["odi"], dependencies=[Depends(login_required)])

templates_dir = settings.get_template_path("app/modulos/odi/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

ALLOWED_UPLOAD_ROLES = [_COLABORADOR, "Colaborador Plus", "Administrador", "Supervisor", _USUARIO]


def get_odi_data_filepath():
    db_dir = getattr(settings, _DB_DIR, None) or DB_DIR
    os.makedirs(db_dir, exist_ok=True)
    return os.path.join(db_dir, "odi_data.json")


def get_odi_storage_dir():
    db_dir = getattr(settings, _DB_DIR, None) or DB_DIR
    storage_dir = os.path.join(db_dir, "odi_files")
    os.makedirs(storage_dir, exist_ok=True)
    return storage_dir


def get_odi_backup_dir():
    db_dir = getattr(settings, _DB_DIR, None) or DB_DIR
    backup_dir = os.path.join(db_dir, "odi_backup")
    os.makedirs(backup_dir, exist_ok=True)
    return backup_dir


def write_odi_log(action: str, user, file_info: dict, extra: str = ""):
    """Registra una línea en odi_log.txt (no visible desde la interfaz)."""
    try:
        who = (getattr(user, _FULL_NAME, None) or getattr(user, _USERNAME, None) or _USUARIO) if user else _USUARIO
        uid = getattr(user, "id", None) if user else None
        line = (
            f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {action} | usuario={who} (id={uid}) | "
            f"archivo={file_info.get('filename')} | id={file_info.get('id')} | "
            f"autor={file_info.get('uploaded_by')} (id={file_info.get('uploader_id')})"
            f"{' | ' + extra if extra else ''}\n"
        )
        with open(os.path.join(get_odi_backup_dir(), "odi_log.txt"), "a", encoding=_UTF_8) as lf:
            lf.write(line)
    except Exception as e:
        print(f"Error escribiendo log ODI: {e}")


def backup_deleted_file(file_info: dict, raw_filepath: str, user):
    """Guarda copia de seguridad (Excel + datos) antes de eliminar, para recuperación."""
    backup_dir = get_odi_backup_dir()
    stamp = datetime.now().strftime(_Y_M_D__H_M_S)
    base = f"{stamp}_{file_info.get('id')}"
    if os.path.exists(raw_filepath):
        shutil.copy2(raw_filepath, os.path.join(backup_dir, f"{base}.xlsx"))
    with open(os.path.join(backup_dir, f"{base}.json"), "w", encoding=_UTF_8) as bf:
        json.dump(file_info, bf, ensure_ascii=False)
    write_odi_log("ELIMINADO", user, file_info, f"copia={base}")


def _group_rows_by_sheet(ruc_map: dict) -> dict:
    sheet_rows: dict[str, list[tuple[str, dict]]] = {}
    for ruc, entries in ruc_map.items():
        for entry in entries:
            sh_name = entry.get(_SHEET, "Hoja1")
            for row in entry.get("rows", []):
                sheet_rows.setdefault(sh_name, []).append((ruc, row))
    return sheet_rows


def _write_rebuilt_sheet(wb, sh_name: str, rows_list: list):
    ws = wb.create_sheet(title=sh_name[:31])
    if not rows_list:
        ws.append(["RUC"])
        return
    extra_keys = []
    seen = set()
    for _, r_dict in rows_list:
        if not isinstance(r_dict, dict):
            continue
        for k in r_dict:
            if k not in seen and k != "RUC":
                seen.add(k)
                extra_keys.append(k)
    ws.append(["RUC"] + extra_keys)
    for ruc_val, r_dict in rows_list:
        ws.append([ruc_val] + [r_dict.get(k, "") for k in extra_keys])


def rebuild_excel_from_data(file_info: dict) -> io.BytesIO:
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = file_info.get(_SHEETS, [])
    sheet_rows = _group_rows_by_sheet(file_info.get(_RUC_MAP, {}))

    if not sheets and not sheet_rows:
        ws = wb.create_sheet(title="Hoja1")
        ws.append(["RUC"])
    else:
        for sh in sheets:
            sh_name = sh.get("name", "Hoja")
            _write_rebuilt_sheet(wb, sh_name, sheet_rows.get(sh_name, []))

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def load_odi_data() -> dict:
    filepath = get_odi_data_filepath()
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding=_UTF_8) as f:
                data = json.load(f)
                if isinstance(data, dict) and _FILES in data:
                    return data
        except Exception as e:
            print(f"Error cargando odi_data.json: {e}")
    return {_FILES: []}


def save_odi_data(data: dict):
    filepath = get_odi_data_filepath()
    with open(filepath, "w", encoding=_UTF_8) as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def build_unified_payload(data: dict, current_user=None) -> dict:
    files = data.get(_FILES, [])
    files_summary = []
    unified_ruc_map: dict[str, list[dict]] = {}
    all_sheets = []

    curr_uid = getattr(current_user, "id", None) if current_user else None
    getattr(current_user, "role", "") if current_user else ""

    for f in files:
        fid = f.get("id")
        fname = f.get(_FILENAME)
        uploader_id = f.get(_UPLOADER_ID)
        uploader_name = f.get(_UPLOADED_BY, "—")

        is_mine = bool(curr_uid is not None and str(uploader_id) == str(curr_uid))
        can_delete = True

        files_summary.append(
            {
                "id": fid,
                _FILENAME: fname,
                _SIZE_KB: f.get(_SIZE_KB, 0),
                _UPLOADED_AT: f.get(_UPLOADED_AT, "—"),
                _UPLOADED_BY: uploader_name,
                _UPLOADER_ID: uploader_id,
                _UPLOADER_ROLE: f.get(_UPLOADER_ROLE, _COLABORADOR),
                _TOTAL_RUCS: f.get(_TOTAL_RUCS, 0),
                "sheets_count": len(f.get(_SHEETS, [])),
                "is_mine": is_mine,
                "can_delete": can_delete,
            }
        )

        for sheet in f.get(_SHEETS, []):
            all_sheets.append(
                {
                    "file_id": fid,
                    "file_name": fname,
                    "name": sheet.get("name"),
                    "ruc_count": sheet.get("ruc_count", 0),
                    "is_mine": is_mine,
                    "uploader_name": uploader_name,
                }
            )

        for ruc, entries in f.get(_RUC_MAP, {}).items():
            for entry in entries:
                unified_ruc_map.setdefault(ruc, []).append(
                    {
                        "file_id": fid,
                        "file_name": fname,
                        _UPLOADER_ID: uploader_id,
                        "uploader_name": uploader_name,
                        _SHEET: entry.get(_SHEET),
                        "rows": entry.get("rows", []),
                    }
                )

    return {
        "success": True,
        "has_data": len(files) > 0,
        "total_files": len(files),
        _TOTAL_RUCS: len(unified_ruc_map),
        "current_user_id": curr_uid,
        _FILES: files_summary,
        "all_sheets": all_sheets,
        "unified_ruc_map": unified_ruc_map,
    }


@router.get("/")
async def odi_home(request: Request):
    user = getattr(request.state, "user", None)
    return templates.TemplateResponse(
        "odi/index.html",
        {
            "request": request,
            "can_upload": True,
            "user_id": user.id if user else None,
            "user_name": (user.full_name if user and user.full_name else getattr(user, _USERNAME, _USUARIO))
            if user
            else _USUARIO,
            "user_role": user.role if user else _COLABORADOR,
        },
    )


@router.get("/api/data")
async def get_stored_data(request: Request):
    user = getattr(request.state, "user", None)
    data = load_odi_data()
    return build_unified_payload(data, current_user=user)


def _uploader_info(user):
    name = (getattr(user, _FULL_NAME, None) or getattr(user, _USERNAME, None)) if user else _USUARIO
    uploader_id = getattr(user, "id", None) or "invitado"
    role = getattr(user, "role", _COLABORADOR) if user else _COLABORADOR
    return name, uploader_id, role


def _save_raw_file(raw_filepath: str, file_id: str, content: bytes):
    try:
        with open(raw_filepath, "wb") as rf:
            rf.write(content)
    except Exception as e:
        print(f"Error al guardar binario de archivo {file_id}: {e}")


def _discard_replaced_file(ef: dict, storage_dir: str, raw_filepath: str, user):
    old_id = ef.get("id")
    if not old_id:
        return
    old_path = os.path.join(storage_dir, f"{old_id}.xlsx")
    try:
        backup_deleted_file(ef, old_path, user)
    except Exception as e:
        print(f"Error respaldando archivo reemplazado: {e}")
    if os.path.exists(old_path) and old_path != raw_filepath:
        with contextlib.suppress(Exception):
            os.remove(old_path)


def _store_payload(existing_files: list, payload: dict, storage_dir: str, raw_filepath: str, user) -> bool:
    """Guarda el archivo sin sobreponer los de otros colaboradores.

    Solo reemplaza si el MISMO colaborador sube un archivo con el MISMO nombre exacto.
    Devuelve True si reemplazó uno existente.
    """
    for idx, ef in enumerate(existing_files):
        if (
            str(ef.get(_UPLOADER_ID)) == str(payload[_UPLOADER_ID])
            and ef.get(_FILENAME, "").lower() == payload[_FILENAME].lower()
        ):
            _discard_replaced_file(ef, storage_dir, raw_filepath, user)
            existing_files[idx] = payload
            return True
    existing_files.append(payload)
    return False


@router.post("/api/upload")
async def upload_excel_files(request: Request, files: list[UploadFile] = File(...)):
    user = getattr(request.state, "user", None)
    data = load_odi_data()
    existing_files = data.get(_FILES, [])
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    uploader_name, uploader_id, uploader_role = _uploader_info(user)

    processed_count = 0
    errors = []

    for i, file in enumerate(files):
        if not file.filename.lower().endswith(".xlsx"):
            errors.append(f"{file.filename}: Solo se admiten archivos .xlsx")
            continue

        try:
            content = await file.read()
            result = analyze_excel(content)
        except Exception as e:
            errors.append(f"{file.filename}: Error al leer archivo ({e})")
            continue

        file_id = f"db_{int(datetime.now().timestamp() * 1000)}_{uploader_id or 0}_{i}"

        # Guardar archivo binario para descargas
        storage_dir = get_odi_storage_dir()
        raw_filepath = os.path.join(storage_dir, f"{file_id}.xlsx")
        _save_raw_file(raw_filepath, file_id, content)

        file_payload = {
            "id": file_id,
            _FILENAME: file.filename,
            _SIZE_KB: round(len(content) / 1024, 1),
            _UPLOADED_AT: now_str,
            _UPLOADED_BY: uploader_name,
            _UPLOADER_ID: uploader_id,
            _UPLOADER_ROLE: uploader_role,
            _SHEETS: result.get(_SHEETS, []),
            _RUC_MAP: result.get(_RUC_MAP, {}),
            _TOTAL_RUCS: result.get(_TOTAL_RUCS, 0),
        }

        replaced = _store_payload(existing_files, file_payload, storage_dir, raw_filepath, user)
        write_odi_log("REEMPLAZADO" if replaced else "SUBIDO", user, file_payload)
        processed_count += 1

    data[_FILES] = existing_files
    try:
        save_odi_data(data)
    except Exception as e:
        print(f"Error al guardar repositorio ODI: {e}")

    payload = build_unified_payload(data, current_user=user)
    payload["processed_count"] = processed_count
    if errors:
        payload["warning"] = " | ".join(errors)
    return payload


@router.delete("/api/files/{file_id}")
async def delete_excel_file(request: Request, file_id: str):
    user = getattr(request.state, "user", None)
    data = load_odi_data()
    original_files = data.get(_FILES, [])
    target_file = next((f for f in original_files if f.get("id") == file_id), None)

    if not target_file:
        return JSONResponse(status_code=404, content={"success": False, _ERROR: "Base de datos no encontrada."})

    updated_files = [f for f in original_files if f.get("id") != file_id]
    data[_FILES] = updated_files

    try:
        storage_dir = get_odi_storage_dir()
        raw_filepath = os.path.join(storage_dir, f"{file_id}.xlsx")
        backup_deleted_file(target_file, raw_filepath, user)
        save_odi_data(data)
        if os.path.exists(raw_filepath):
            os.remove(raw_filepath)
    except Exception as e:
        print(f"Error al guardar tras eliminar base de datos en ODI: {e}")

    return build_unified_payload(data, current_user=user)


@router.get("/api/files/{file_id}/download")
async def download_excel_file(file_id: str):
    data = load_odi_data()
    file_info = None
    for f in data.get(_FILES, []):
        if f.get("id") == file_id:
            file_info = f
            break

    if not file_info:
        return JSONResponse(status_code=404, content={_ERROR: "Base de datos no encontrada."})

    filename = file_info.get(_FILENAME, "Base_Datos.xlsx")
    storage_dir = get_odi_storage_dir()
    raw_path = os.path.join(storage_dir, f"{file_id}.xlsx")

    if os.path.exists(raw_path):
        with open(raw_path, "rb") as rf:
            file_bytes = rf.read()
    else:
        # Reconstrucción de respaldo si el archivo se subió antes de habilitar almacenamiento físico
        rebuilt = rebuild_excel_from_data(file_info)
        file_bytes = rebuilt.getvalue()

    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            _CONTENT_DISPOSITION: f'attachment; filename="{filename}"',
            _ACCESS_CONTROL_EXPOSE_HEADER: _CONTENT_DISPOSITION,
        },
    )


@router.post("/api/export/excel")
async def export_classification_excel(request: Request):
    user = getattr(request.state, "user", None)
    user_name = getattr(user, _FULL_NAME, None) or getattr(user, _USERNAME, _USUARIO)
    body = await request.json()
    results = body.get("results", [])

    if not results:
        return JSONResponse(status_code=400, content={_ERROR: "No hay resultados para exportar."})

    try:
        stream = generate_excel_export(results, user_name=user_name)
        now_ts = datetime.now().strftime(_Y_M_D__H_M_S)
        filename = f"Clasificacion_ODI_{now_ts}.xlsx"
        return Response(
            content=stream.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                _CONTENT_DISPOSITION: f'attachment; filename="{filename}"',
                _ACCESS_CONTROL_EXPOSE_HEADER: _CONTENT_DISPOSITION,
            },
        )
    except Exception as e:
        print(f"Error generando exportación Excel ODI: {e}")
        return JSONResponse(status_code=500, content={_ERROR: f"Error al generar Excel: {e!s}"})


@router.post("/api/export/pdf")
async def export_classification_pdf(request: Request):
    user = getattr(request.state, "user", None)
    user_name = getattr(user, _FULL_NAME, None) or getattr(user, _USERNAME, _USUARIO)
    body = await request.json()
    results = body.get("results", [])

    if not results:
        return JSONResponse(status_code=400, content={_ERROR: "No hay resultados para exportar."})

    try:
        pdf_bytes = generate_pdf_export(results, user_name=user_name)
        now_ts = datetime.now().strftime(_Y_M_D__H_M_S)
        filename = f"Reporte_ODI_{now_ts}.pdf"
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                _CONTENT_DISPOSITION: f'attachment; filename="{filename}"',
                _ACCESS_CONTROL_EXPOSE_HEADER: _CONTENT_DISPOSITION,
            },
        )
    except Exception as e:
        print(f"Error generando exportación PDF ODI: {e}")
        return JSONResponse(status_code=500, content={_ERROR: f"Error al generar PDF: {e!s}"})
