# -*- coding: utf-8 -*-

import pandas as pd


def cargar_excel(path):
    """Carga un Excel intentando varios motores para soportar .xls y .xlsx de forma robusta."""

    motores_a_probar = []

    # 1. Deducir motor prioritario por extensión
    if path.lower().endswith(".xls"):
        motores_a_probar.append("xlrd")  # Prioridad para .xls
        motores_a_probar.append("openpyxl")  # Fallback por si acaso es un xlsx renombrado
    elif path.lower().endswith(".xlsx"):
        motores_a_probar.append("openpyxl")  # Prioridad para .xlsx
        motores_a_probar.append("xlrd")  # Fallback por si acaso es un xls renombrado
    else:
        # Sin extensión clara, probamos openpyxl primero (estándar moderno)
        motores_a_probar = ["openpyxl", "xlrd"]

    errores = []

    # 2. Intentar con los motores en orden
    for engine in motores_a_probar:
        try:
            df = pd.read_excel(path, engine=engine)
            df.columns = df.columns.str.strip().str.lower()
            return df
        except Exception as e:
            errores.append(f"Engine '{engine}': {str(e)}")

    # 3. Último intento: Automático (sin especificar engine, pandas decide)
    try:
        df = pd.read_excel(path)
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        errores.append(f"Auto: {str(e)}")

    # Si llegamos aquí, nada funcionó
    errores_str = " | ".join(errores)
    raise ValueError(f"No se pudo leer el archivo Excel. Detalles técnicos: {errores_str}")


def fmt_fecha(x):
    try:
        val = pd.to_datetime(x, dayfirst=True, errors="coerce")
        if pd.isna(val):
            return ""
        return val.strftime("%d/%m/%Y")
    except Exception:
        return ""


CODIGOS_RRAF = ("080201", "080300", "052309", "053302")


def _validar_tipo_archivo(deuda, validar_tipo):
    """Valida que el archivo de deuda corresponda al módulo (FRACC / RRAF)."""
    if not (validar_tipo and "tipo" in deuda.columns):
        return
    # Valores de la columna 'tipo' en un solo texto para buscar palabras clave
    valores_presentes = " ".join(deuda["tipo"].dropna().astype(str).unique()).upper()

    if validar_tipo == "FRACC" and "RRAF" in valores_presentes:
        raise ValueError(
            "Error Crítico: Ha subido un archivo de REFINANCIAMIENTO (RRAF) en el módulo de Fraccionamiento. Por favor verifique su archivo."
        )
    if validar_tipo == "RRAF" and any(k in valores_presentes for k in ("ART. 36", "ART 36", "ART.36")):
        raise ValueError(
            "Error Crítico: Ha subido un archivo de FRACCIONAMIENTO (Art. 36) en el módulo de Refinanciamiento. Por favor verifique su archivo."
        )


def _normalizar_codigos(deuda, pagos):
    if "cod_tri" in deuda.columns:
        deuda["cod_tri"] = deuda["cod_tri"].astype(str).str.zfill(6)
    if "cod_tributo" in pagos.columns:
        pagos["cod_tributo"] = pagos["cod_tributo"].astype(str).str.zfill(6)
    # Reemplazo cod_tri_aso -> cod_tri
    if "cod_tri_aso" in deuda.columns and "cod_tri" not in deuda.columns:
        deuda.rename(columns={"cod_tri_aso": "cod_tri"}, inplace=True)


def _validar_columnas(deuda, pagos):
    req_deuda = ["num_doc_rin", "num_ruc", "cod_tri", "per_doc"]
    req_pagos = ["ruc", "doc_pago", "doc_que_paga", "fec_pago", "periodo", "cod_tributo", "mto_pago"]
    for col in req_deuda:
        if col not in deuda.columns:
            raise ValueError(f"Falta columna '{col}' en DEUDA ACOGIDA")
    for col in req_pagos:
        if col not in pagos.columns:
            raise ValueError(f"Falta columna '{col}' en PAGOS")


def _filtrar_pagos_por_ria(deuda, pagos, ria_num):
    """Acota los pagos a la RIA indicada. Devuelve (pagos, fecha_maestra)."""
    fecha_maestra = pd.NaT
    if not ria_num:
        return pagos, fecha_maestra

    deuda_ria = deuda[deuda["num_doc_rin"].astype(str) == str(ria_num)]
    if deuda_ria.empty:
        pagos = pagos[pagos["doc_que_paga"].astype(str) == str(ria_num)]
        if pagos.empty:
            raise ValueError(f"No se encontraron pagos asociados a la RIA {ria_num}")
        return pagos, fecha_maestra

    ruc_asociado = deuda_ria["num_ruc"].iloc[0]
    fecha_maestra = deuda_ria["fec_aprobacion_dt"].iloc[0]
    pagos["fec_pago_dt"] = pd.to_datetime(pagos["fec_pago"], dayfirst=True, errors="coerce")
    pagos = pagos[pagos["ruc"].astype(str) == str(ruc_asociado)]
    if pd.notna(fecha_maestra):
        pagos = pagos[pagos["fec_pago_dt"] >= fecha_maestra]
    if pagos.empty:
        raise ValueError(f"No existen pagos para el RUC asociado a la RIA {ria_num}")
    return pagos, fecha_maestra


def _calcular_monto_acogido(deuda, ria_num):
    if ria_num and "mto_aco_sol" in deuda.columns:
        temp_deuda = deuda[deuda["num_doc_rin"].astype(str) == str(ria_num)]
        return pd.to_numeric(temp_deuda["mto_aco_sol"], errors="coerce").fillna(0).sum()
    return 0.0


def _construir_base(deuda, pagos, fecha_maestra):
    """Une los pagos con las fechas de aprobación de RIA, RRAF y deuda acogida."""
    base = pagos.copy()
    base["fec_aprobacion_maestra"] = fecha_maestra

    deuda_ria = deuda[["num_doc_rin", "fec_aprobacion_dt"]].drop_duplicates()
    base = base.merge(
        deuda_ria.rename(columns={"num_doc_rin": "ria_doc", "fec_aprobacion_dt": "fec_aprobacion_ria"}),
        left_on="doc_que_paga",
        right_on="ria_doc",
        how="left",
    )

    # RRAF cabecera (080300)
    rraf_cab = deuda[["num_doc_rin", "fec_aprobacion_dt"]].drop_duplicates()
    base = base.merge(
        rraf_cab.rename(columns={"num_doc_rin": "rraf_cab", "fec_aprobacion_dt": "fec_rraf"}),
        left_on="doc_que_paga",
        right_on="rraf_cab",
        how="left",
    )

    # Deuda acogida detalle (RRAF)
    rraf_det = deuda[deuda["cod_tri"].astype(str).str.zfill(6) == "080201"][["num_doc", "fec_aprobacion_dt"]]
    base = base.merge(
        rraf_det.rename(columns={"num_doc": "rraf_det", "fec_aprobacion_dt": "fec_rraf_det"}),
        left_on="doc_que_paga",
        right_on="rraf_det",
        how="left",
    )

    # Código-periodo (RIA normal)
    deuda_codper = deuda[["num_ruc", "cod_tri", "per_doc", "fec_aprobacion_dt"]].drop_duplicates()
    base = base.merge(
        deuda_codper.rename(
            columns={
                "num_ruc": "ruc_d",
                "cod_tri": "cod_tributo_d",
                "per_doc": "periodo_d",
                "fec_aprobacion_dt": "fec_aprobacion_deuda",
            }
        ),
        left_on=["ruc", "cod_tributo", "periodo"],
        right_on=["ruc_d", "cod_tributo_d", "periodo_d"],
        how="left",
    )

    base["fec_pago_dt"] = pd.to_datetime(base["fec_pago"], dayfirst=True, errors="coerce")
    base["fec_aprobacion_deuda"] = (
        base["fec_aprobacion_deuda"].fillna(base["fec_aprobacion_ria"]).fillna(base["fec_aprobacion_maestra"])
    )
    return base


def _clasificar_pago(row):
    cod = str(row.get("cod_tributo", "")).zfill(6)

    if pd.notna(row.get("fec_rraf")) and cod in ("080300", "052309", "053302"):
        return "PAGO RRAF"
    if pd.notna(row.get("fec_rraf_det")) and cod == "080201":
        return "PAGO DEUDA ACOGIDA (RRAF)"
    if pd.notna(row.get("fec_aprobacion_ria")) and cod in ("080201", "052106", "053105"):
        return "PAGO FRACCIONAMIENTO"
    if row.get("cod_tributo_d") == row.get("cod_tributo") and row.get("periodo_d") == row.get("periodo"):
        return "PAGO DEUDA ACOGIDA (RIA)"
    return "PAGO SUELTO"


def _aplicar_filtros_correctivos(base):
    cods = base["cod_tributo"].astype(str).str.zfill(6)
    base = base[~((base["TipoCruce"] == "PAGO SUELTO") & (~cods.isin(CODIGOS_RRAF)))]
    if base["TipoCruce"].str.contains("RRAF", na=False).any():
        base = base[base["cod_tributo"].astype(str).str.zfill(6).isin(CODIGOS_RRAF)]
    return base


def _filtro_final_ria(base, ria_num):
    doc = str(ria_num)
    filtro_directos = base[base["doc_que_paga"].astype(str) == doc]
    filtro_deuda = base[base["TipoCruce"].str.contains("DEUDA ACOGIDA", na=False)]
    filtro_suelto = base[
        (base["cod_tributo"].astype(str).str.zfill(6).isin(CODIGOS_RRAF))
        & (base["doc_pago"].astype(str) == base["doc_que_paga"].astype(str))
    ]
    return pd.concat([filtro_directos, filtro_deuda, filtro_suelto]).drop_duplicates()


def _formatear_salida(base):
    base["fec_aprobacion"] = (
        base["fec_aprobacion_deuda"]
        .fillna(base["fec_rraf"])
        .fillna(base["fec_rraf_det"])
        .fillna(base["fec_aprobacion_maestra"])
        .map(fmt_fecha)
    )
    base["fec_pago"] = base["fec_pago_dt"].map(fmt_fecha)

    base["__ord"] = pd.to_datetime(base["fec_pago"], dayfirst=True, errors="coerce")
    base = base.sort_values("__ord").drop(columns="__ord")

    campos = ["TipoCruce", "ruc", "doc_que_paga", "fec_aprobacion", "fec_pago", "periodo", "cod_tributo", "mto_pago"]
    for c in campos:
        if c not in base.columns:
            base[c] = ""
    return base[campos].reset_index(drop=True)


def ejecutar_cruce(path_deuda, path_pagos, ria_num="", validar_tipo=None):
    deuda = cargar_excel(path_deuda)
    pagos = cargar_excel(path_pagos)

    deuda.columns = deuda.columns.str.strip().str.lower()
    pagos.columns = pagos.columns.str.strip().str.lower()

    _validar_tipo_archivo(deuda, validar_tipo)
    _normalizar_codigos(deuda, pagos)
    _validar_columnas(deuda, pagos)

    # Fecha de aprobación (columna 22 del Excel de deuda)
    try:
        deuda["fec_aprobacion_dt"] = pd.to_datetime(deuda.iloc[:, 21], dayfirst=True, errors="coerce")
    except Exception:
        deuda["fec_aprobacion_dt"] = pd.NaT

    deuda["es_rraf"] = False  # placeholder sin uso

    pagos, fecha_maestra = _filtrar_pagos_por_ria(deuda, pagos, ria_num)
    monto_acogido = _calcular_monto_acogido(deuda, ria_num)

    base = _construir_base(deuda, pagos, fecha_maestra)
    base["TipoCruce"] = base.apply(_clasificar_pago, axis=1)
    base = _aplicar_filtros_correctivos(base)
    if ria_num:
        base = _filtro_final_ria(base, ria_num)

    return _formatear_salida(base), monto_acogido


def result_summary(df, monto_acogido=0):
    pagos_deu = len(df[df["TipoCruce"].str.contains("DEUDA ACOGIDA", na=False)])
    pagos_fracc = len(df[df["TipoCruce"].str.contains("FRACCIONAMIENTO", na=False)])
    pagos_suelto = len(df[df["TipoCruce"].str.contains("SUELTO", na=False)])
    pagos_rraf = len(df[df["TipoCruce"].str.contains("RRAF", na=False)])
    total = len(df)
    monto_total = pd.to_numeric(df["mto_pago"], errors="coerce").fillna(0).sum()

    return {
        "pagos_deuda": int(pagos_deu),
        "pagos_fracc": int(pagos_fracc),
        "pagos_sueltos": int(pagos_suelto),
        "pagos_rraf": int(pagos_rraf),
        "total": int(total),
        "monto_total": float(monto_total),
        "monto_acogido": float(monto_acogido),
    }


def ejecutar_cruce_fracc(path_deuda, path_pagos, ria_num=""):
    df, monto = ejecutar_cruce(path_deuda, path_pagos, ria_num, validar_tipo="FRACC")
    df = df[df["TipoCruce"].isin(["PAGO FRACCIONAMIENTO", "PAGO DEUDA ACOGIDA (RIA)", "PAGO SUELTO"])].reset_index(
        drop=True
    )
    return df, monto


def ejecutar_cruce_rraf(path_deuda, path_pagos, ria_num=""):
    df, monto = ejecutar_cruce(path_deuda, path_pagos, ria_num, validar_tipo="RRAF")
    df = df[df["TipoCruce"].isin(["PAGO RRAF", "PAGO DEUDA ACOGIDA (RRAF)", "PAGO SUELTO"])].reset_index(drop=True)
    return df, monto
