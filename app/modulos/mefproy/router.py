"""
MEFProy Router - Endpoints para consulta y automatización de Proyectos PTE del MEF.
"""
from fastapi import APIRouter, Request, Query, Response, Depends, HTTPException, Body
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional, Dict, Any

from .service import MEFProyService, UNIDADES_EJECUTORAS_PRESET
from app.modulos.auth.dependencies import login_required
from app.core.config import settings

router = APIRouter(
    prefix="/mefproy",
    tags=["mefproy"],
    dependencies=[Depends(login_required)]
)

templates_dir = settings.get_template_path("app/modulos/mefproy/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

service = MEFProyService()


class ExportarMEFRequest(BaseModel):
    proyectos: List[Dict[str, Any]]
    id_ue: str = "818"
    periodo: str = "2026"
    tipo_vista: str = "ejecucion"


@router.get("/")
async def mefproy_home(
    request: Request,
    id_ue: str = Query("818"),
    periodo: str = Query("2026")
):
    """Renderiza el dashboard interactivo de Proyectos MEF."""
    presets = service.obtener_presets_ue()
    periodos = service.obtener_periodos("Ejecucion")
    
    return templates.TemplateResponse(
        "mefproy/index.html",
        {
            "request": request,
            "presets": presets,
            "periodos": periodos,
            "default_ue": id_ue,
            "default_periodo": periodo
        }
    )


@router.get("/api/periodos")
async def get_periodos(tipo: str = Query("Ejecucion")):
    """Retorna los años fiscales disponibles en el MEF."""
    try:
        data = service.obtener_periodos(tipo)
        return {"status": "success", "data": data}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})


@router.get("/api/presets")
async def get_presets():
    """Retorna las unidades ejecutoras preconfiguradas."""
    return {"status": "success", "data": service.obtener_presets_ue()}


@router.get("/api/ejecucion")
def get_ejecucion(
    id_ue: str = Query("818"),
    periodo: str = Query("2026"),
    page_size: int = Query(50),
    page: int = Query(1),
    sort_col: str = Query("codDGPP"),
    sort_order: str = Query("asc"),
    q: str = Query(""),
    nom_proyecto: str = Query(""),
    cod_dgpp: str = Query(""),
    cod_snip: str = Query("")
):
    """Consulta la ejecución presupuestal de proyectos en tiempo real desde el MEF."""
    try:
        final_nom = nom_proyecto.strip()
        final_dgpp = cod_dgpp.strip()
        final_snip = cod_snip.strip()
        query_val = q.strip()

        if query_val and not (final_nom or final_dgpp or final_snip):
            if query_val.isdigit():
                final_dgpp = query_val
            else:
                final_nom = query_val

        result = service.obtener_ejecucion(
            id_ue=id_ue,
            periodo=periodo,
            page_size=page_size,
            page_number=page,
            sort_column=sort_col,
            sort_order=sort_order,
            nom_proyecto=final_nom,
            cod_dgpp=final_dgpp,
            cod_snip=final_snip
        )

        if query_val and query_val.isdigit() and len(result.get("proyectos", [])) == 0 and not (nom_proyecto or cod_dgpp or cod_snip):
            result = service.obtener_ejecucion(
                id_ue=id_ue,
                periodo=periodo,
                page_size=page_size,
                page_number=page,
                sort_column=sort_col,
                sort_order=sort_order,
                nom_proyecto=query_val,
                cod_dgpp="",
                cod_snip=query_val
            )

        return result
    except Exception as e:
        status_code = 504 if "Timeout" in str(e) or "timed out" in str(e).lower() else 500
        return JSONResponse(status_code=status_code, content={"status": "error", "message": str(e)})


@router.get("/api/banco")
def get_banco(
    id_ue: str = Query("818"),
    periodo: str = Query("2026"),
    page_size: int = Query(50),
    page: int = Query(1),
    sort_col: str = Query("codDGPP"),
    sort_order: str = Query("asc")
):
    """Consulta el Banco de Proyectos (DGPI) con fechas de registro y viabilidad."""
    try:
        result = service.obtener_banco_proyectos(
            id_ue=id_ue,
            periodo=periodo,
            page_size=page_size,
            page_number=page,
            sort_column=sort_col,
            sort_order=sort_order
        )
        return result
    except Exception as e:
        status_code = 504 if "Timeout" in str(e) or "timed out" in str(e).lower() else 500
        return JSONResponse(status_code=status_code, content={"status": "error", "message": str(e)})


@router.get("/api/ranking")
def get_ranking(
    id_ue: str = Query("818"),
    periodo: str = Query("2026")
):
    """Obtiene el Top de Proyectos con mayor PIM para la UE seleccionada."""
    try:
        data = service.obtener_ranking(id_ue=id_ue, periodo=periodo)
        return {"status": "success", "data": data}
    except Exception as e:
        status_code = 504 if "Timeout" in str(e) or "timed out" in str(e).lower() else 500
        return JSONResponse(status_code=status_code, content={"status": "error", "message": str(e)})


@router.get("/api/detalle")
def get_detalle(
    id_ue: str = Query("818"),
    cod_proyecto: str = Query(...)
):
    """Obtiene el histórico multianual y trimestral de un proyecto."""
    try:
        data = service.obtener_detalle_proyecto(id_ue=id_ue, cod_proyecto=cod_proyecto)
        return {"status": "success", "data": data}
    except Exception as e:
        status_code = 504 if "Timeout" in str(e) or "timed out" in str(e).lower() else 500
        return JSONResponse(status_code=status_code, content={"status": "error", "message": str(e)})


@router.post("/api/exportar")
def exportar_mef(data: ExportarMEFRequest):
    """Genera archivo Excel (.xlsx) estructurado y estilizado para descarga."""
    try:
        xlsx_bytes = service.generar_excel_xlsx(
            proyectos=data.proyectos,
            id_ue=data.id_ue,
            periodo=data.periodo,
            tipo_vista=data.tipo_vista
        )
        filename = f"Proyectos_MEF_{data.tipo_vista.upper()}_UE_{data.id_ue}_{data.periodo}.xlsx"
        return Response(
            content=xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

