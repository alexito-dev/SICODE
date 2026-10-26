from fastapi import APIRouter, Request, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
import os
import shutil
import pandas as pd
from datetime import datetime
import tempfile
import uuid

from app.core.config import settings
from app.modulos.piloto.logic import ejecutar_cruce_fracc, ejecutar_cruce_rraf, result_summary

# Textos repetidos
_XLSX = ".xlsx"

router = APIRouter(prefix="/piloto", tags=["piloto"])

# Configuración de plantillas
templates_dir = settings.get_template_path("app/modulos/piloto/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates_global = settings.get_template_path("app/templates")

templates = Jinja2Templates(directory=[templates_dir, templates_base, templates_global])

# In-memory storage for results (for download)
# In a real app, this should be a DB or cache
results_cache = {}


@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("piloto/index.html", {"request": request})


@router.post("/process")
async def process_cruce(
    modulo: str = Form(...),
    ria_num: str = Form(""),
    deuda_file: UploadFile = File(...),
    pagos_file: UploadFile = File(...),
):
    try:
        # Create temp files for logic processing
        with tempfile.NamedTemporaryFile(delete=False, suffix=_XLSX) as tmp_deuda:
            shutil.copyfileobj(deuda_file.file, tmp_deuda)
            path_deuda = tmp_deuda.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=_XLSX) as tmp_pagos:
            shutil.copyfileobj(pagos_file.file, tmp_pagos)
            path_pagos = tmp_pagos.name

        if modulo == "FRACC":
            df, monto_acogido = ejecutar_cruce_fracc(path_deuda, path_pagos, ria_num)
        else:
            df, monto_acogido = ejecutar_cruce_rraf(path_deuda, path_pagos, ria_num)

        summary = result_summary(df, monto_acogido)

        # Convert DF to dict for JSON response
        data_records = df.where(pd.notnull(df), None).to_dict(orient="records")

        # Cache for download
        file_id = str(uuid.uuid4())
        results_cache[file_id] = df

        # Cleanup temp files
        os.remove(path_deuda)
        os.remove(path_pagos)

        return JSONResponse({"status": "success", "summary": summary, "data": data_records, "file_id": file_id})

    except Exception as e:
        return JSONResponse({"status": "error", "message": str(e)})


@router.get("/download/{file_id}")
async def download_results(file_id: str):
    if file_id not in results_cache:
        raise HTTPException(status_code=404, detail="File result not found")

    df = results_cache[file_id]

    with tempfile.NamedTemporaryFile(delete=False, suffix=_XLSX) as tmp:
        output_path = tmp.name
        df.to_excel(output_path, index=False)

    return FileResponse(
        output_path,
        filename=f"RESULTADO_PILOTO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=None,  # shutil will handle deletion if needed or use a task
    )
