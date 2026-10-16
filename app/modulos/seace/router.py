from fastapi import APIRouter, Request, Depends, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from app.core.config import settings
from app.modulos.auth.dependencies import login_required
from app.modulos.seace.seace_service import process_rucs, generate_excel_bytes, fetch_contracts_for_ruc
import pandas as pd
import io
import os
import tempfile
from datetime import datetime

router = APIRouter(
    prefix="/seace", 
    tags=["Seace"],
    dependencies=[Depends(login_required)]
)

# Configuración de plantillas
templates_dir = settings.get_template_path("app/modulos/seace/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates_global = settings.get_template_path("app/templates")

templates = Jinja2Templates(directory=[templates_dir, templates_base, templates_global])

@router.get("/", response_class=HTMLResponse)
async def seace_home(request: Request):
    # Registrar acceso
    log_user = getattr(request.app.state, "log_user", None)
    if log_user:
        user_name = request.state.user.nombre if hasattr(request.state, "user") else "Usuario"
        log_user(user_name, "Entró a SEACE", request.client.host)
    return templates.TemplateResponse("seace/index.html", {"request": request})

@router.post("/process")
async def process_seace_data(
    request: Request,
    rucs: str = Form(None),
    file: UploadFile = File(None),        # Archivo con lista de RUCs
    file_cruce: UploadFile = File(None),  # El Directorio (Opcional para cruce automático)
    start_date: str = Form(None)
):
    rucs_to_process = []
    directorio_df = None
    
    # 0. Procesar Directorio si viene para extraer RUCs (si no hay otros)
    if file_cruce:
        try:
            content_cruce = await file_cruce.read()
            directorio_df = pd.read_excel(io.BytesIO(content_cruce))
            # Reiniciar puntero para poder leerlo luego en el cruce
            await file_cruce.seek(0)
            
            # Si no pasaron RUCs explícitamente, intentamos sacarlos del Directorio
            if not rucs and not file:
                # Prioridad columna num_doc_ident o RUC
                possible_cols = ["num_doc_ident", "RUC", "ruc", "num_doc"]
                target_col = next((c for c in possible_cols if c in directorio_df.columns), None)
                if target_col:
                    rucs_to_process.extend(directorio_df[target_col].astype(str).tolist())
                else:
                    rucs_to_process.extend(directorio_df.iloc[:, 0].astype(str).tolist())
        except Exception as e:
            print(f"Error procesando directorio inicial: {e}")

    # 1. Obtener RUCs del Textarea
    if rucs:
        rucs_to_process.extend(rucs.replace('\r', '').split('\n'))
        
    # 2. Obtener RUCs del archivo de lista
    if file:
        try:
            content = await file.read()
            if file.filename.endswith('.xlsx') or file.filename.endswith('.xls'):
                df = pd.read_excel(io.BytesIO(content))
                possible_cols = [col for col in df.columns if 'RUC' in str(col).upper()]
                if possible_cols:
                    rucs_to_process.extend(df[possible_cols[0]].astype(str).tolist())
                else:
                    rucs_to_process.extend(df.iloc[:, 0].astype(str).tolist())
            else:
                 text_content = content.decode('utf-8')
                 rucs_to_process.extend(text_content.replace('\r', '').split('\n'))
        except Exception as e:
            return JSONResponse(status_code=400, content={"status": "error", "message": f"Error leyendo lista de RUCs: {str(e)}"})

    # Limpiar lista
    rucs_to_process = [r.strip() for r in rucs_to_process if r.strip().isdigit() and len(r.strip()) == 11]
    rucs_to_process = list(set(rucs_to_process))

    if not rucs_to_process:
         return JSONResponse(status_code=400, content={"status": "error", "message": "No se encontraron RUCs válidos para consultar."})

    # PROCESAR EN SEACE
    try:
        # Registrar actividad de procesamiento
        log_user = getattr(request.app.state, "log_user", None)
        if log_user:
            user_name = request.state.user.nombre if hasattr(request.state, "user") else "Usuario"
            log_user(user_name, f"Buscando {len(rucs_to_process)} RUCs", request.client.host)
            
        print(f"[DEBUG] RUCs a procesar: {len(rucs_to_process)} -> {rucs_to_process[:5]}")
        print(f"[DEBUG] start_date: '{start_date}'")
        processing_result = await process_rucs(rucs_to_process, start_date)
        results = processing_result["results"]
        success_rucs = processing_result["success_rucs"]
        failed_rucs = processing_result["failed_rucs"]
        print(f"[DEBUG] Exitosos: {len(success_rucs)}, Fallidos: {len(failed_rucs)}, Contratos: {len(results)}")
        if failed_rucs:
            print(f"[DEBUG] Primer fallido: {failed_rucs[0]}")
        
        # Generar DataFrames Básicos (df_simple tiene 9 campos, df_datos tiene 10)
        df_datos, df_fallidos, df_simple = generate_excel_bytes(results, failed_rucs)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"REPORTE_SEACE_{timestamp}.xlsx"
        report_path = os.path.join(tempfile.gettempdir(), report_filename)
        
        # Guardar Reporte Normal (Usamos df_simple como pidió el usuario)
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            df_simple.to_excel(writer, sheet_name="Datos", index=False)
            df_fallidos.to_excel(writer, sheet_name="Fallidos", index=False)
        
        response_data = {
            "status": "success",
            "count": len(results),
            "success_count": len(success_rucs),
            "success_rucs": success_rucs,
            "failed_count": len(failed_rucs),
            "failed_rucs": failed_rucs,
            "download_url": f"/seace/download/{report_filename}"
        }

        # --- LÓGICA DE CRUCE PROFESIONAL (Si se subió el Directorio) ---
        if directorio_df is not None:
            try:
                directorio_df["num_doc_ident"] = directorio_df["num_doc_ident"].astype(str)
                cols_extraer = ["DIRECTORIO", "RUC", "RAZON SOCIAL", "DEUDA TOTAL", "DEUDA COBRANZA COACTIVA", "num_doc_ident"]
                
                if all(c in directorio_df.columns for c in cols_extraer):
                    excel1_reducido = directorio_df[cols_extraer].copy()
                    excel1_reducido = excel1_reducido.rename(columns={"RUC": "RUC_CONSORCIO"})

                    df_simple["RUC_STR"] = df_simple["RUC"].astype(str)
                    cruce_datos = df_simple.merge(excel1_reducido, left_on="RUC_STR", right_on="num_doc_ident", how="left")
                    cruce_datos = cruce_datos.drop(columns=["num_doc_ident", "RUC_STR"])
                    cruce_datos = cruce_datos.rename(columns={"RUC": "RUC_PROVEEDOR"})

                    columnas_prefix = ["DIRECTORIO", "RUC_CONSORCIO", "RAZON SOCIAL", "DEUDA TOTAL", "DEUDA COBRANZA COACTIVA"]
                    columnas_resto = [c for c in cruce_datos.columns if c not in columnas_prefix]
                    cruce_datos = cruce_datos[columnas_prefix + columnas_resto]

                    df_fallidos = df_fallidos.rename(columns={"Motivo": "MOTIVO"})
                    df_fallidos["RUC_STR"] = df_fallidos["RUC"].astype(str)
                    cruce_fallidos = df_fallidos.merge(excel1_reducido, left_on="RUC_STR", right_on="num_doc_ident", how="left")
                    cruce_fallidos = cruce_fallidos.drop(columns=["num_doc_ident", "RUC_STR"])
                    cruce_fallidos = cruce_fallidos.rename(columns={"RUC": "RUC_PROVEEDOR"})
                    
                    columnas_fallidos = columnas_prefix + ["RUC_PROVEEDOR", "MOTIVO"]
                    cruce_fallidos = cruce_fallidos[columnas_fallidos]

                    cruce_filename = f"REPORTE_CON_CRUCE_{timestamp}.xlsx"
                    cruce_path = os.path.join(tempfile.gettempdir(), cruce_filename)
                    with pd.ExcelWriter(cruce_path, engine="openpyxl") as writer:
                        cruce_datos.to_excel(writer, sheet_name="Cruce_Datos", index=False)
                        cruce_fallidos.to_excel(writer, sheet_name="Cruce_Fallidos", index=False)
                    
                    response_data["cruce_url"] = f"/seace/download/{cruce_filename}"
                    response_data["success_count"] = int(cruce_datos["RUC_PROVEEDOR"].nunique())
                    response_data["failed_count"] = len(cruce_fallidos)
                    response_data["count"] = len(cruce_datos)
                    
                    new_failed_list = []
                    for _, row in cruce_fallidos.iterrows():
                        consorcio_nombre = str(row.get("RAZON SOCIAL", "N/A"))
                        new_failed_list.append({
                            "ruc": str(row.get("RUC_PROVEEDOR", "")),
                            "reason": f"[{consorcio_nombre}] {row.get('MOTIVO', '')}"
                        })
                    response_data["failed_rucs"] = new_failed_list

                else:
                    response_data["cruce_error"] = "El Directorio no tiene todas las columnas necesarias para el cruce."
            except Exception as e_cruce:
                response_data["cruce_error"] = f"Error en el cruce automático: {str(e_cruce)}"

        return response_data
        
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})

@router.post("/cruce-profesional")
async def cruce_profesional(
    request: Request,
    file_directorio: UploadFile = File(...),
    file_reporte: UploadFile = File(...)
):
    try:
        # Registrar actividad
        log_user = getattr(request.app.state, "log_user", None)
        if log_user:
            user_name = request.state.user.nombre if hasattr(request.state, "user") else "Usuario"
            log_user(user_name, "Inició Cruce Profesional", request.client.host)
            
        # 1. Leer Archivos
        content1 = await file_directorio.read()
        content2 = await file_reporte.read()
        
        excel1 = pd.read_excel(io.BytesIO(content1))
        excel1["num_doc_ident"] = excel1["num_doc_ident"].astype(str)
        
        hojas_excel2 = pd.read_excel(io.BytesIO(content2), sheet_name=None)
        
        if "Datos" not in hojas_excel2 or "Fallidos" not in hojas_excel2:
            return JSONResponse(status_code=400, content={"status": "error", "message": "El reporte SEACE debe tener las hojas 'Datos' y 'Fallidos'"})

        hoja_datos = hojas_excel2["Datos"]
        hoja_datos["RUC"] = hoja_datos["RUC"].astype(str)

        hoja_fallidos = hojas_excel2["Fallidos"]
        hoja_fallidos["RUC"] = hoja_fallidos["RUC"].astype(str)

        # 2. Filtrar columnas del Directorio
        cols_necesarias = ["num_doc_ident", "DIRECTORIO", "RUC", "RAZON SOCIAL", "DEUDA TOTAL", "DEUDA COBRANZA COACTIVA"]
        
        for col in cols_necesarias:
            if col not in excel1.columns:
                return JSONResponse(status_code=400, content={"status": "error", "message": f"Falta columna '{col}' en el Excel de Directorio"})

        excel1_reducido = excel1[cols_necesarias].copy()
        excel1_reducido = excel1_reducido.rename(columns={"RUC": "RUC_CONSORCIO"})

        # 3. Cruce Hoja DATOS
        cruce_datos = hoja_datos.merge(
            excel1_reducido,
            left_on="RUC",
            right_on="num_doc_ident",
            how="left"
        )
        cruce_datos = cruce_datos.drop(columns=["num_doc_ident"])
        cruce_datos = cruce_datos.rename(columns={"RUC": "RUC_PROVEEDOR"})

        columnas_prefix = ["DIRECTORIO", "RUC_CONSORCIO", "RAZON SOCIAL", "DEUDA TOTAL", "DEUDA COBRANZA COACTIVA"]
        columnas_resto = [col for col in cruce_datos.columns if col not in columnas_prefix]
        cruce_datos = cruce_datos[columnas_prefix + columnas_resto]

        # 4. Cruce Hoja FALLIDOS
        cruce_fallidos = hoja_fallidos.merge(
            excel1_reducido,
            left_on="RUC",
            right_on="num_doc_ident",
            how="left"
        )
        cruce_fallidos = cruce_fallidos.drop(columns=["num_doc_ident"])
        cruce_fallidos = cruce_fallidos.rename(columns={"RUC": "RUC_PROVEEDOR"})
        
        columnas_resto_fallidos = [col for col in cruce_fallidos.columns if col not in columnas_prefix]
        cruce_fallidos = cruce_fallidos[columnas_prefix + columnas_resto_fallidos]

        # 5. Guardar Resultado
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"CRUCE_PROFESIONAL_{timestamp}.xlsx"
        filepath = os.path.join(tempfile.gettempdir(), filename)
        
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            cruce_datos.to_excel(writer, sheet_name="Cruce_Datos", index=False)
            cruce_fallidos.to_excel(writer, sheet_name="Cruce_Fallidos", index=False)

        return {
            "status": "success",
            "count": len(cruce_datos),
            "success_count": int(cruce_datos["RUC_PROVEEDOR"].nunique()),
            "failed_count": len(cruce_fallidos),
            "download_url": f"/seace/download/{filename}"
        }

    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": f"Error en el cruce: {str(e)}"})

@router.get("/download/{filename}")
async def download_report(filename: str):
    filepath = os.path.join(tempfile.gettempdir(), filename)
    if os.path.exists(filepath):
        return FileResponse(filepath, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return JSONResponse(status_code=404, content={"message": "Archivo no encontrado"})
