from fastapi import APIRouter, Request, Form, Depends, HTTPException
from fastapi.templating import Jinja2Templates
from fastapi.responses import JSONResponse, FileResponse
import os
from app.modulos.auth.dependencies import login_required
from app.modulos.signa.logic import procesar_carpeta

router = APIRouter(prefix="/signa", tags=["signa"], dependencies=[Depends(login_required)])

from app.core.config import settings

templates_dir = settings.get_template_path("app/modulos/signa/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])


@router.get("/")
async def signa_home(request: Request):
    return templates.TemplateResponse("signa/index.html", {"request": request})


@router.post("/verify")
async def verify_signature_process(folder_path: str = Form(...)):
    if not os.path.isdir(folder_path):
        return JSONResponse(
            {"status": "error", "message": f"La ruta especificada no es una carpeta válida: {folder_path}"},
            status_code=400,
        )

    try:
        # Run processing logic
        result = procesar_carpeta(folder_path)

        # Add the folder_path to result for download link construction
        result["folder_path"] = folder_path

        return JSONResponse(result)

    except Exception as e:
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@router.get("/download")
async def download_report(path: str, filename: str):
    file_path = os.path.join(path, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
