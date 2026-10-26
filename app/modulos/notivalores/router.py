"""
NotiValores Router - Endpoints para simulación de consulta y notificación de Valores SUNAT.
"""

from fastapi import APIRouter, Request, Response, Depends, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import json
import csv
import io

from .processor import NotiValoresProcessor
from app.modulos.auth.dependencies import login_required
from app.core.config import settings

router = APIRouter(prefix="/notivalores", tags=["notivalores"], dependencies=[Depends(login_required)])

templates_dir = settings.get_template_path("app/modulos/notivalores/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

processor = NotiValoresProcessor()


class SimularRequest(BaseModel):
    rucs: List[str]
    canal: Optional[str] = "sol"
    incluir_sin_deuda: Optional[bool] = False


class CedulaRequest(BaseModel):
    valor: Dict[str, Any]


class ExportarRequest(BaseModel):
    valores: List[Dict[str, Any]]


@router.get("/")
async def notivalores_home(request: Request):
    """Renderiza el dashboard interactivo del módulo NotiValores."""
    presets = processor.obtener_presets()
    return templates.TemplateResponse("notivalores/index.html", {"request": request, "presets": presets})


@router.get("/api/presets")
async def get_presets():
    """Retorna los RUCs preconfigurados para pruebas rápidas."""
    return {"status": "success", "data": processor.obtener_presets()}


@router.post("/api/simular")
async def simular_notificacion(data: SimularRequest):
    """
    Streaming SSE que ejecuta el proceso simulado de búsqueda y notificación
    de valores SUNAT para la lista de RUCs dada.
    """

    async def event_generator():
        async for update in processor.simular_procesamiento_rucs(
            rucs_raw=data.rucs, canal=data.canal, incluir_sin_deuda=data.incluir_sin_deuda
        ):
            yield f"data: {json.dumps(update, ensure_ascii=False)}\n\n"

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


@router.post("/api/cedula")
async def ver_cedula(data: CedulaRequest):
    """Genera la Cédula de Notificación en formato HTML oficial."""
    try:
        html_content = processor.generar_html_cedula(data.valor)
        return HTMLResponse(content=html_content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@router.post("/api/exportar")
async def exportar_resultados(data: ExportarRequest):
    """Exporta la lista de valores simulados a CSV para Excel."""
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";", quoting=csv.QUOTE_MINIMAL)

    # Encabezados
    writer.writerow(
        [
            "ID Valor",
            "RUC",
            "Razón Social",
            "Tipo de Valor",
            "Nivel Riesgo",
            "N° Documento",
            "Período",
            "Tributo",
            "Monto Deuda (S/)",
            "Estado",
            "Fecha Emisión",
            "Canal Notificación",
            "Notificado",
        ]
    )

    for item in data.valores:
        writer.writerow(
            [
                item.get("id", ""),
                item.get("ruc", ""),
                item.get("razon_social", ""),
                item.get("tipo_valor", ""),
                item.get("riesgo", ""),
                item.get("num_documento", ""),
                item.get("periodo", ""),
                item.get("tributo", ""),
                f"{item.get('monto_deuda', 0):.2f}",
                item.get("estado", ""),
                item.get("fecha_emision", ""),
                item.get("canal_notificacion", ""),
                "SI" if item.get("notificado") else "NO",
            ]
        )

    csv_bytes = output.getvalue().encode("utf-8-sig")  # UTF-8 con BOM para Excel
    output.close()

    return Response(
        content=csv_bytes,
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=Reporte_NotiValores_SUNAT.csv"},
    )
