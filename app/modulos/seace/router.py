from fastapi import APIRouter, Request, Depends, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from app.core.config import settings
from app.modulos.auth.dependencies import login_required
from app.modulos.seace.seace_service import process_rucs, generate_excel_bytes
import pandas as pd
import io
import os
import tempfile
from datetime import datetime

# Textos repetidos
_COUNT = "count"
_DATOS = "Datos"
_DEUDA_COBRANZA_COACTIVA = "DEUDA COBRANZA COACTIVA"
_DEUDA_TOTAL = "DEUDA TOTAL"
_DIRECTORIO = "DIRECTORIO"
_ERROR = "error"
_FAILED_COUNT = "failed_count"
_FAILED_RUCS = "failed_rucs"
_FALLIDOS = "Fallidos"
_LOG_USER = "log_user"
_MESSAGE = "message"
_NUM_DOC_IDENT = "num_doc_ident"
_OPENPYXL = "openpyxl"
_RAZON_SOCIAL = "RAZON SOCIAL"
_RUC_CONSORCIO = "RUC_CONSORCIO"
_RUC_PROVEEDOR = "RUC_PROVEEDOR"
_RUC_STR = "RUC_STR"
_STATUS = "status"
_SUCCESS_COUNT = "success_count"
_USUARIO = "Usuario"

router = APIRouter(prefix="/seace", tags=["Seace"], dependencies=[Depends(login_required)])

# Configuración de plantillas
templates_dir = settings.get_template_path("app/modulos/seace/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates_global = settings.get_template_path("app/templates")

templates = Jinja2Templates(directory=[templates_dir, templates_base, templates_global])


@router.get("/", response_class=HTMLResponse)
async def seace_home(request: Request):
    # Registrar acceso
    log_user = getattr(request.app.state, _LOG_USER, None)
    if log_user:
        user_name = request.state.user.nombre if hasattr(request.state, "user") else _USUARIO
        log_user(user_name, "Entró a SEACE", request.client.host)
    return templates.TemplateResponse("seace/index.html", {"request": request})


_COLS_DIRECTORIO = [_DIRECTORIO, "RUC", _RAZON_SOCIAL, _DEUDA_TOTAL, _DEUDA_COBRANZA_COACTIVA, _NUM_DOC_IDENT]
_COLS_PREFIJO_CRUCE = [_DIRECTORIO, _RUC_CONSORCIO, _RAZON_SOCIAL, _DEUDA_TOTAL, _DEUDA_COBRANZA_COACTIVA]


async def _leer_directorio(file_cruce, rucs, file, rucs_to_process):
    """Lee el Directorio y, si no hay otra fuente de RUCs, los toma de ahí."""
    try:
        content_cruce = await file_cruce.read()
        directorio_df = pd.read_excel(io.BytesIO(content_cruce))
        # Reiniciar puntero para poder leerlo luego en el cruce
        await file_cruce.seek(0)

        if not rucs and not file:
            # Prioridad columna num_doc_ident o RUC
            possible_cols = [_NUM_DOC_IDENT, "RUC", "ruc", "num_doc"]
            target_col = next((c for c in possible_cols if c in directorio_df.columns), None)
            if target_col:
                rucs_to_process.extend(directorio_df[target_col].astype(str).tolist())
            else:
                rucs_to_process.extend(directorio_df.iloc[:, 0].astype(str).tolist())
        return directorio_df
    except Exception as e:
        print(f"Error procesando directorio inicial: {e}")
        return None


async def _rucs_desde_archivo(file):
    content = await file.read()
    if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(content))
        possible_cols = [col for col in df.columns if "RUC" in str(col).upper()]
        columna = df[possible_cols[0]] if possible_cols else df.iloc[:, 0]
        return columna.astype(str).tolist()
    return content.decode("utf-8").replace("\r", "").split("\n")


def _registrar_busqueda(request, cantidad):
    log_user = getattr(request.app.state, _LOG_USER, None)
    if log_user:
        user_name = request.state.user.nombre if hasattr(request.state, "user") else _USUARIO
        log_user(user_name, f"Buscando {cantidad} RUCs", request.client.host)


def _cruzar_con_directorio(directorio_df, df_simple, df_fallidos, timestamp, response_data):
    """Cruza datos y fallidos con el Directorio y completa response_data."""
    directorio_df[_NUM_DOC_IDENT] = directorio_df[_NUM_DOC_IDENT].astype(str)
    if not all(c in directorio_df.columns for c in _COLS_DIRECTORIO):
        response_data["cruce_error"] = "El Directorio no tiene todas las columnas necesarias para el cruce."
        return

    excel1_reducido = directorio_df[_COLS_DIRECTORIO].copy().rename(columns={"RUC": _RUC_CONSORCIO})

    df_simple[_RUC_STR] = df_simple["RUC"].astype(str)
    cruce_datos = df_simple.merge(excel1_reducido, left_on=_RUC_STR, right_on=_NUM_DOC_IDENT, how="left")
    cruce_datos = cruce_datos.drop(columns=[_NUM_DOC_IDENT, _RUC_STR]).rename(columns={"RUC": _RUC_PROVEEDOR})
    columnas_resto = [c for c in cruce_datos.columns if c not in _COLS_PREFIJO_CRUCE]
    cruce_datos = cruce_datos[_COLS_PREFIJO_CRUCE + columnas_resto]

    df_fallidos = df_fallidos.rename(columns={"Motivo": "MOTIVO"})
    df_fallidos[_RUC_STR] = df_fallidos["RUC"].astype(str)
    cruce_fallidos = df_fallidos.merge(excel1_reducido, left_on=_RUC_STR, right_on=_NUM_DOC_IDENT, how="left")
    cruce_fallidos = cruce_fallidos.drop(columns=[_NUM_DOC_IDENT, _RUC_STR]).rename(
        columns={"RUC": _RUC_PROVEEDOR}
    )
    cruce_fallidos = cruce_fallidos[_COLS_PREFIJO_CRUCE + [_RUC_PROVEEDOR, "MOTIVO"]]

    cruce_filename = f"REPORTE_CON_CRUCE_{timestamp}.xlsx"
    cruce_path = os.path.join(tempfile.gettempdir(), cruce_filename)
    with pd.ExcelWriter(cruce_path, engine=_OPENPYXL) as writer:
        cruce_datos.to_excel(writer, sheet_name="Cruce_Datos", index=False)
        cruce_fallidos.to_excel(writer, sheet_name="Cruce_Fallidos", index=False)

    response_data["cruce_url"] = f"/seace/download/{cruce_filename}"
    response_data[_SUCCESS_COUNT] = int(cruce_datos[_RUC_PROVEEDOR].nunique())
    response_data[_FAILED_COUNT] = len(cruce_fallidos)
    response_data[_COUNT] = len(cruce_datos)
    response_data[_FAILED_RUCS] = [
        {
            "ruc": str(row.get(_RUC_PROVEEDOR, "")),
            "reason": f"[{row.get('RAZON SOCIAL', 'N/A')!s}] {row.get('MOTIVO', '')}",
        }
        for _, row in cruce_fallidos.iterrows()
    ]


@router.post("/process")
async def process_seace_data(
    request: Request,
    rucs: str = Form(None),
    file: UploadFile = File(None),  # Archivo con lista de RUCs
    file_cruce: UploadFile = File(None),  # El Directorio (Opcional para cruce automático)
    start_date: str = Form(None),
):
    rucs_to_process = []
    directorio_df = None

    # 0. Directorio (opcional): puede aportar los RUCs si no hay otra fuente
    if file_cruce:
        directorio_df = await _leer_directorio(file_cruce, rucs, file, rucs_to_process)

    # 1. RUCs del Textarea
    if rucs:
        rucs_to_process.extend(rucs.replace("\r", "").split("\n"))

    # 2. RUCs del archivo de lista
    if file:
        try:
            rucs_to_process.extend(await _rucs_desde_archivo(file))
        except Exception as e:
            return JSONResponse(
                status_code=400, content={_STATUS: _ERROR, _MESSAGE: f"Error leyendo lista de RUCs: {e!s}"}
            )

    # Limpiar lista
    rucs_to_process = [r.strip() for r in rucs_to_process if r.strip().isdigit() and len(r.strip()) == 11]
    rucs_to_process = list(set(rucs_to_process))

    if not rucs_to_process:
        return JSONResponse(
            status_code=400, content={_STATUS: _ERROR, _MESSAGE: "No se encontraron RUCs válidos para consultar."}
        )

    # PROCESAR EN SEACE
    try:
        _registrar_busqueda(request, len(rucs_to_process))

        print(f"[DEBUG] RUCs a procesar: {len(rucs_to_process)} -> {rucs_to_process[:5]}")
        print(f"[DEBUG] start_date: '{start_date}'")
        processing_result = await process_rucs(rucs_to_process, start_date)
        results = processing_result["results"]
        success_rucs = processing_result["success_rucs"]
        failed_rucs = processing_result[_FAILED_RUCS]
        print(f"[DEBUG] Exitosos: {len(success_rucs)}, Fallidos: {len(failed_rucs)}, Contratos: {len(results)}")
        if failed_rucs:
            print(f"[DEBUG] Primer fallido: {failed_rucs[0]}")

        # DataFrames básicos (df_simple tiene 9 campos, df_datos tiene 10)
        _df_datos, df_fallidos, df_simple = generate_excel_bytes(results, failed_rucs)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"REPORTE_SEACE_{timestamp}.xlsx"
        report_path = os.path.join(tempfile.gettempdir(), report_filename)

        with pd.ExcelWriter(report_path, engine=_OPENPYXL) as writer:
            df_simple.to_excel(writer, sheet_name=_DATOS, index=False)
            df_fallidos.to_excel(writer, sheet_name=_FALLIDOS, index=False)

        response_data = {
            _STATUS: "success",
            _COUNT: len(results),
            _SUCCESS_COUNT: len(success_rucs),
            "success_rucs": success_rucs,
            _FAILED_COUNT: len(failed_rucs),
            _FAILED_RUCS: failed_rucs,
            "download_url": f"/seace/download/{report_filename}",
        }

        if directorio_df is not None:
            try:
                _cruzar_con_directorio(directorio_df, df_simple, df_fallidos, timestamp, response_data)
            except Exception as e_cruce:
                response_data["cruce_error"] = f"Error en el cruce automático: {e_cruce!s}"

        return response_data

    except Exception as e:
        return JSONResponse(status_code=500, content={_STATUS: _ERROR, _MESSAGE: str(e)})


@router.post("/cruce-profesional")
async def cruce_profesional(
    request: Request, file_directorio: UploadFile = File(...), file_reporte: UploadFile = File(...)
):
    try:
        # Registrar actividad
        log_user = getattr(request.app.state, _LOG_USER, None)
        if log_user:
            user_name = request.state.user.nombre if hasattr(request.state, "user") else _USUARIO
            log_user(user_name, "Inició Cruce Profesional", request.client.host)

        # 1. Leer Archivos
        content1 = await file_directorio.read()
        content2 = await file_reporte.read()

        excel1 = pd.read_excel(io.BytesIO(content1))
        excel1[_NUM_DOC_IDENT] = excel1[_NUM_DOC_IDENT].astype(str)

        hojas_excel2 = pd.read_excel(io.BytesIO(content2), sheet_name=None)

        if _DATOS not in hojas_excel2 or _FALLIDOS not in hojas_excel2:
            return JSONResponse(
                status_code=400,
                content={_STATUS: _ERROR, _MESSAGE: "El reporte SEACE debe tener las hojas 'Datos' y 'Fallidos'"},
            )

        hoja_datos = hojas_excel2[_DATOS]
        hoja_datos["RUC"] = hoja_datos["RUC"].astype(str)

        hoja_fallidos = hojas_excel2[_FALLIDOS]
        hoja_fallidos["RUC"] = hoja_fallidos["RUC"].astype(str)

        # 2. Filtrar columnas del Directorio
        cols_necesarias = [
            _NUM_DOC_IDENT,
            _DIRECTORIO,
            "RUC",
            _RAZON_SOCIAL,
            _DEUDA_TOTAL,
            _DEUDA_COBRANZA_COACTIVA,
        ]

        for col in cols_necesarias:
            if col not in excel1.columns:
                return JSONResponse(
                    status_code=400,
                    content={_STATUS: _ERROR, _MESSAGE: f"Falta columna '{col}' en el Excel de Directorio"},
                )

        excel1_reducido = excel1[cols_necesarias].copy()
        excel1_reducido = excel1_reducido.rename(columns={"RUC": _RUC_CONSORCIO})

        # 3. Cruce Hoja DATOS
        cruce_datos = hoja_datos.merge(excel1_reducido, left_on="RUC", right_on=_NUM_DOC_IDENT, how="left")
        cruce_datos = cruce_datos.drop(columns=[_NUM_DOC_IDENT])
        cruce_datos = cruce_datos.rename(columns={"RUC": _RUC_PROVEEDOR})

        columnas_prefix = [_DIRECTORIO, _RUC_CONSORCIO, _RAZON_SOCIAL, _DEUDA_TOTAL, _DEUDA_COBRANZA_COACTIVA]
        columnas_resto = [col for col in cruce_datos.columns if col not in columnas_prefix]
        cruce_datos = cruce_datos[columnas_prefix + columnas_resto]

        # 4. Cruce Hoja FALLIDOS
        cruce_fallidos = hoja_fallidos.merge(excel1_reducido, left_on="RUC", right_on=_NUM_DOC_IDENT, how="left")
        cruce_fallidos = cruce_fallidos.drop(columns=[_NUM_DOC_IDENT])
        cruce_fallidos = cruce_fallidos.rename(columns={"RUC": _RUC_PROVEEDOR})

        columnas_resto_fallidos = [col for col in cruce_fallidos.columns if col not in columnas_prefix]
        cruce_fallidos = cruce_fallidos[columnas_prefix + columnas_resto_fallidos]

        # 5. Guardar Resultado
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"CRUCE_PROFESIONAL_{timestamp}.xlsx"
        filepath = os.path.join(tempfile.gettempdir(), filename)

        with pd.ExcelWriter(filepath, engine=_OPENPYXL) as writer:
            cruce_datos.to_excel(writer, sheet_name="Cruce_Datos", index=False)
            cruce_fallidos.to_excel(writer, sheet_name="Cruce_Fallidos", index=False)

        return {
            _STATUS: "success",
            _COUNT: len(cruce_datos),
            _SUCCESS_COUNT: int(cruce_datos[_RUC_PROVEEDOR].nunique()),
            _FAILED_COUNT: len(cruce_fallidos),
            "download_url": f"/seace/download/{filename}",
        }

    except Exception as e:
        return JSONResponse(status_code=500, content={_STATUS: _ERROR, _MESSAGE: f"Error en el cruce: {e!s}"})


@router.get("/download/{filename}")
async def download_report(filename: str):
    filepath = os.path.join(tempfile.gettempdir(), filename)
    if os.path.exists(filepath):
        return FileResponse(
            filepath, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return JSONResponse(status_code=404, content={_MESSAGE: "Archivo no encontrado"})
