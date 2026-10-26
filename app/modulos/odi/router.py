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

router = APIRouter(
    prefix="/odi",
    tags=["odi"],
    dependencies=[Depends(login_required)]
)

templates_dir = settings.get_template_path("app/modulos/odi/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

ALLOWED_UPLOAD_ROLES = ["Colaborador", "Colaborador Plus", "Administrador", "Supervisor", "Usuario"]


def get_odi_data_filepath():
    db_dir = getattr(settings, "DB_DIR", None) or DB_DIR
    os.makedirs(db_dir, exist_ok=True)
    return os.path.join(db_dir, "odi_data.json")


def get_odi_storage_dir():
    db_dir = getattr(settings, "DB_DIR", None) or DB_DIR
    storage_dir = os.path.join(db_dir, "odi_files")
    os.makedirs(storage_dir, exist_ok=True)
    return storage_dir


def get_odi_backup_dir():
    db_dir = getattr(settings, "DB_DIR", None) or DB_DIR
    backup_dir = os.path.join(db_dir, "odi_backup")
    os.makedirs(backup_dir, exist_ok=True)
    return backup_dir


def write_odi_log(action: str, user, file_info: dict, extra: str = ""):
    """Registra una línea en odi_log.txt (no visible desde la interfaz)."""
    try:
        who = (getattr(user, "full_name", None) or getattr(user, "username", None) or "Usuario") if user else "Usuario"
        uid = getattr(user, "id", None) if user else None
        line = (f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {action} | usuario={who} (id={uid}) | "
                f"archivo={file_info.get('filename')} | id={file_info.get('id')} | "
                f"autor={file_info.get('uploaded_by')} (id={file_info.get('uploader_id')})"
                f"{' | ' + extra if extra else ''}\n")
        with open(os.path.join(get_odi_backup_dir(), "odi_log.txt"), "a", encoding="utf-8") as lf:
            lf.write(line)
    except Exception as e:
        print(f"Error escribiendo log ODI: {e}")


def backup_deleted_file(file_info: dict, raw_filepath: str, user):
    """Guarda copia de seguridad (Excel + datos) antes de eliminar, para recuperación."""
    backup_dir = get_odi_backup_dir()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"{stamp}_{file_info.get('id')}"
    if os.path.exists(raw_filepath):
        shutil.copy2(raw_filepath, os.path.join(backup_dir, f"{base}.xlsx"))
    with open(os.path.join(backup_dir, f"{base}.json"), "w", encoding="utf-8") as bf:
        json.dump(file_info, bf, ensure_ascii=False)
    write_odi_log("ELIMINADO", user, file_info, f"copia={base}")


def rebuild_excel_from_data(file_info: dict) -> io.BytesIO:
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheets = file_info.get("sheets", [])
    ruc_map = file_info.get("ruc_map", {})

    sheet_rows: dict[str, list[tuple[str, dict]]] = {}
    for ruc, entries in ruc_map.items():
        for entry in entries:
            sh_name = entry.get("sheet", "Hoja1")
            for row in entry.get("rows", []):
                sheet_rows.setdefault(sh_name, []).append((ruc, row))

    if not sheets and not sheet_rows:
        ws = wb.create_sheet(title="Hoja1")
        ws.append(["RUC"])
    else:
        for sh in sheets:
            sh_name = sh.get("name", "Hoja")
            ws = wb.create_sheet(title=sh_name[:31])
            rows_list = sheet_rows.get(sh_name, [])
            if rows_list:
                extra_keys = []
                seen = set()
                for _, r_dict in rows_list:
                    if isinstance(r_dict, dict):
                        for k in r_dict.keys():
                            if k not in seen and k != "RUC":
                                seen.add(k)
                                extra_keys.append(k)
                headers = ["RUC"] + extra_keys
                ws.append(headers)
                for ruc_val, r_dict in rows_list:
                    row_vals = [ruc_val] + [r_dict.get(k, "") for k in extra_keys]
                    ws.append(row_vals)
            else:
                ws.append(["RUC"])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def load_odi_data() -> dict:
    filepath = get_odi_data_filepath()
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "files" in data:
                    return data
        except Exception as e:
            print(f"Error cargando odi_data.json: {e}")
    return {"files": []}


def save_odi_data(data: dict):
    filepath = get_odi_data_filepath()
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def build_unified_payload(data: dict, current_user=None) -> dict:
    files = data.get("files", [])
    files_summary = []
    unified_ruc_map: dict[str, list[dict]] = {}
    all_sheets = []

    curr_uid = getattr(current_user, "id", None) if current_user else None
    curr_role = getattr(current_user, "role", "") if current_user else ""
    is_admin_or_super = curr_role in ["Administrador", "Supervisor"]

    for f in files:
        fid = f.get("id")
        fname = f.get("filename")
        uploader_id = f.get("uploader_id")
        uploader_name = f.get("uploaded_by", "—")

        is_mine = bool(curr_uid is not None and str(uploader_id) == str(curr_uid))
        can_delete = True

        files_summary.append({
            "id": fid,
            "filename": fname,
            "size_kb": f.get("size_kb", 0),
            "uploaded_at": f.get("uploaded_at", "—"),
            "uploaded_by": uploader_name,
            "uploader_id": uploader_id,
            "uploader_role": f.get("uploader_role", "Colaborador"),
            "total_rucs": f.get("total_rucs", 0),
            "sheets_count": len(f.get("sheets", [])),
            "is_mine": is_mine,
            "can_delete": can_delete
        })

        for sheet in f.get("sheets", []):
            all_sheets.append({
                "file_id": fid,
                "file_name": fname,
                "name": sheet.get("name"),
                "ruc_count": sheet.get("ruc_count", 0),
                "is_mine": is_mine,
                "uploader_name": uploader_name
            })

        for ruc, entries in f.get("ruc_map", {}).items():
            for entry in entries:
                unified_ruc_map.setdefault(ruc, []).append({
                    "file_id": fid,
                    "file_name": fname,
                    "uploader_id": uploader_id,
                    "uploader_name": uploader_name,
                    "sheet": entry.get("sheet"),
                    "rows": entry.get("rows", [])
                })

    return {
        "success": True,
        "has_data": len(files) > 0,
        "total_files": len(files),
        "total_rucs": len(unified_ruc_map),
        "current_user_id": curr_uid,
        "files": files_summary,
        "all_sheets": all_sheets,
        "unified_ruc_map": unified_ruc_map
    }


@router.get("/")
async def odi_home(request: Request):
    user = getattr(request.state, "user", None)
    return templates.TemplateResponse("odi/index.html", {
        "request": request,
        "can_upload": True,
        "user_id": user.id if user else None,
        "user_name": (user.full_name if user and user.full_name else getattr(user, "username", "Usuario")) if user else "Usuario",
        "user_role": user.role if user else "Colaborador"
    })


@router.get("/api/data")
async def get_stored_data(request: Request):
    user = getattr(request.state, "user", None)
    data = load_odi_data()
    return build_unified_payload(data, current_user=user)


@router.post("/api/upload")
async def upload_excel_files(request: Request, files: list[UploadFile] = File(...)):
    user = getattr(request.state, "user", None)
    data = load_odi_data()
    existing_files = data.get("files", [])
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    uploader_name = (getattr(user, "full_name", None) or getattr(user, "username", None)) if user else "Usuario"
    uploader_id = getattr(user, "id", None) or "invitado"
    uploader_role = getattr(user, "role", "Colaborador") if user else "Colaborador"

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
        try:
            with open(raw_filepath, "wb") as rf:
                rf.write(content)
        except Exception as e:
            print(f"Error al guardar binario de archivo {file_id}: {e}")

        file_payload = {
            "id": file_id,
            "filename": file.filename,
            "size_kb": round(len(content) / 1024, 1),
            "uploaded_at": now_str,
            "uploaded_by": uploader_name,
            "uploader_id": uploader_id,
            "uploader_role": uploader_role,
            "sheets": result.get("sheets", []),
            "ruc_map": result.get("ruc_map", {}),
            "total_rucs": result.get("total_rucs", 0)
        }

        # NO sobreponer archivos entre distintos colaboradores:
        # Solo reemplazar si el MISMO colaborador sube un archivo con el MISMO nombre exacto.
        replaced = False
        for idx, ef in enumerate(existing_files):
            if str(ef.get("uploader_id")) == str(uploader_id) and ef.get("filename", "").lower() == file.filename.lower():
                old_id = ef.get("id")
                if old_id:
                    old_path = os.path.join(storage_dir, f"{old_id}.xlsx")
                    try:
                        backup_deleted_file(ef, old_path, user)
                    except Exception as e:
                        print(f"Error respaldando archivo reemplazado: {e}")
                    if os.path.exists(old_path) and old_path != raw_filepath:
                        try:
                            os.remove(old_path)
                        except Exception:
                            pass
                existing_files[idx] = file_payload
                replaced = True
                break
        if not replaced:
            existing_files.append(file_payload)

        write_odi_log("REEMPLAZADO" if replaced else "SUBIDO", user, file_payload)
        processed_count += 1

    data["files"] = existing_files
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
    original_files = data.get("files", [])
    target_file = next((f for f in original_files if f.get("id") == file_id), None)

    if not target_file:
        return JSONResponse(status_code=404, content={"success": False, "error": "Base de datos no encontrada."})

    updated_files = [f for f in original_files if f.get("id") != file_id]
    data["files"] = updated_files

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
async def download_excel_file(request: Request, file_id: str):
    data = load_odi_data()
    file_info = None
    for f in data.get("files", []):
        if f.get("id") == file_id:
            file_info = f
            break

    if not file_info:
        return JSONResponse(status_code=404, content={"error": "Base de datos no encontrada."})

    filename = file_info.get("filename", "Base_Datos.xlsx")
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
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
    )


@router.post("/api/export/excel")

async def export_classification_excel(request: Request):
    user = getattr(request.state, "user", None)
    user_name = getattr(user, "full_name", None) or getattr(user, "username", "Usuario")
    body = await request.json()
    results = body.get("results", [])

    if not results:
        return JSONResponse(status_code=400, content={"error": "No hay resultados para exportar."})

    try:
        stream = generate_excel_export(results, user_name=user_name)
        now_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Clasificacion_ODI_{now_ts}.xlsx"
        return Response(
            content=stream.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    except Exception as e:
        print(f"Error generando exportación Excel ODI: {e}")
        return JSONResponse(status_code=500, content={"error": f"Error al generar Excel: {str(e)}"})


@router.post("/api/export/pdf")
async def export_classification_pdf(request: Request):
    user = getattr(request.state, "user", None)
    user_name = getattr(user, "full_name", None) or getattr(user, "username", "Usuario")
    body = await request.json()
    results = body.get("results", [])

    if not results:
        return JSONResponse(status_code=400, content={"error": "No hay resultados para exportar."})

    try:
        pdf_bytes = generate_pdf_export(results, user_name=user_name)
        now_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Reporte_ODI_{now_ts}.pdf"
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    except Exception as e:
        print(f"Error generando exportación PDF ODI: {e}")
        return JSONResponse(status_code=500, content={"error": f"Error al generar PDF: {str(e)}"})

