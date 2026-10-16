"""
CSV2Excel Router - Endpoints FastAPI para carga masiva, previsualización y conversión de CSVs a Excel.
"""
from fastapi import APIRouter, Request, UploadFile, File, Form, Depends, HTTPException, Response
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from typing import List, Optional
import io
import json
import zipfile

from .processor import CSV2ExcelProcessor, THEMES
from app.modulos.auth.dependencies import login_required
from app.core.config import settings

router = APIRouter(
    prefix="/csv2excel",
    tags=["csv2excel"],
    dependencies=[Depends(login_required)]
)

templates_dir = settings.get_template_path("app/modulos/csv2excel/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

processor = CSV2ExcelProcessor()


@router.get("/")
async def csv2excel_home(request: Request):
    """Renderiza la interfaz interactiva del módulo CSV2Excel."""
    return templates.TemplateResponse(
        "csv2excel/index.html",
        {
            "request": request,
            "themes": THEMES
        }
    )


@router.post("/api/preview")
async def preview_batch(files: List[UploadFile] = File(...)):
    """
    Analiza un lote de archivos subidos (CSV/TXT/TSV) y retorna vista previa de metadatos,
    columnas e inferencia de tipos sin procesar aún la descarga final.
    """
    if not files:
        raise HTTPException(status_code=400, detail="No se enviaron archivos en la solicitud.")

    batch_preview = []

    for file_item in files:
        try:
            content = await file_item.read()
            if file_item.filename.endswith(".zip"):
                # Si se sube un ZIP, extraer y analizar cada CSV adentro
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for z_name in zf.namelist():
                        if z_name.lower().endswith((".csv", ".txt", ".tsv")) and not z_name.startswith("__MACOSX"):
                            z_bytes = zf.read(z_name)
                            p_info = processor.parse_csv_file(z_bytes, filename=z_name)
                            del p_info["df"]  # No enviar DataFrame pesado en JSON
                            batch_preview.append(p_info)
            else:
                p_info = processor.parse_csv_file(content, filename=file_item.filename)
                del p_info["df"]
                batch_preview.append(p_info)
        except Exception as e:
            batch_preview.append({
                "filename": file_item.filename,
                "error": str(e)
            })

    return {"status": "success", "total_files": len(batch_preview), "data": batch_preview}


@router.post("/api/convert")
async def convert_batch(
    files: List[UploadFile] = File(...),
    output_mode: str = Form("multi_sheet"),  # "multi_sheet", "merged_single", "individual_zip"
    theme_key: str = Form("excel_green"),
    preserve_leading_zeros: bool = Form(True),
    add_totals_row: bool = Form(True),
    auto_fit_columns: bool = Form(True),
    freeze_header: bool = Form(True),
    add_index_sheet: bool = Form(True)
):
    """
    Procesa el lote de archivos subidos según las configuraciones y retorna el archivo .xlsx o .zip final.
    """
    if not files:
        raise HTTPException(status_code=400, detail="Debes seleccionar al menos un archivo CSV.")

    parsed_batch = []

    for file_item in files:
        try:
            content = await file_item.read()
            if file_item.filename.endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for z_name in zf.namelist():
                        if z_name.lower().endswith((".csv", ".txt", ".tsv")) and not z_name.startswith("__MACOSX"):
                            z_bytes = zf.read(z_name)
                            p_info = processor.parse_csv_file(z_bytes, filename=z_name)
                            parsed_batch.append(p_info)
            else:
                p_info = processor.parse_csv_file(content, filename=file_item.filename)
                parsed_batch.append(p_info)
        except Exception as e:
            continue

    if not parsed_batch:
        raise HTTPException(status_code=400, detail="No se pudieron procesar los archivos CSV proporcionados.")

    out_bytes, out_filename, media_type = processor.generate_excel_workbook(
        files_data=parsed_batch,
        output_mode=output_mode,
        theme_key=theme_key,
        preserve_leading_zeros=preserve_leading_zeros,
        add_totals_row=add_totals_row,
        auto_fit_columns=auto_fit_columns,
        freeze_header=freeze_header,
        add_index_sheet=add_index_sheet
    )

    return Response(
        content=out_bytes,
        media_type=media_type,
        headers={
            "Content-Disposition": f'attachment; filename="{out_filename}"'
        }
    )


@router.get("/api/demo")
async def download_demo_batch():
    """
    Genera y ofrece la descarga inmediata de un paquete de prueba ZIP con 3 archivos CSV
    de prueba típicos (SUNAT Notificaciones, Presupuesto MEF, Registro de Compras).
    """
    csv_sunat = (
        "RUC;RAZON_SOCIAL;TIPO_DOCUMENTO;NUMERO_VALOR;PERIODO;MONTO_DEUDA;FECHA_EMISION;ESTADO\n"
        "20100123456;CORPORACION INKA S.A.C.;RESOLUCION DE DETERMINACION;011-003-0098765;2024-01;14500.50;15/01/2024;PENDIENTE\n"
        "20500987654;INVERSIONES DEL SUR E.I.R.L.;ORDEN DE PAGO;023-001-0043210;2024-02;890.00;20/02/2024;NOTIFICADO\n"
        "08123456;PERSONA NATURAL CON NEGOCIO;RESOLUCION DE MULTA;055-002-0011223;2023-12;3250.00;10/12/2023;EN COBRANZA\n"
    )

    csv_mef = (
        "CODIGO_UNICO,PROYECTO_INVERSION,ENTIDAD,DEPARTAMENTO,PIM_MONTO,DEVENGADO_MONTO,AVANCE_PCT\n"
        "2345123,CONSTRUCCION DE CENTRO DE SALUD CHIVAY,GOBIERNOS LOCALES,AREQUIPA,4500000.00,3800000.00,84.4\n"
        "0123456,MEJORAMIENTO DE VIAS EN DISTRITO OLIVOS,MUNICIPALIDAD LOS OLIVOS,LIMA,1200000.00,1150000.00,95.8\n"
        "0098712,AMPLIACION RED DE AGUA POTABLE CUSCO,MINISTERIO DE VIVIENDA,CUSCO,8900000.00,6200000.00,69.6\n"
    )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Lote_SUNAT_Valores_Demo.csv", csv_sunat)
        zf.writestr("Lote_MEF_Proyectos_Demo.csv", csv_mef)

    return Response(
        content=zip_buf.getvalue(),
        media_type="application/zip",
        headers={
            "Content-Disposition": 'attachment; filename="Ejemplo_Lote_CSVs_Microsoft.zip"'
        }
    )
