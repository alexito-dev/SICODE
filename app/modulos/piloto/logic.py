# -*- coding: utf-8 -*-

import pandas as pd

# Textos repetidos
_COD_TRI = "cod_tri"
_COD_TRIBUTO = "cod_tributo"
_COD_TRIBUTO_D = "cod_tributo_d"
_COERCE = "coerce"
_DOC_QUE_PAGA = "doc_que_paga"
_FEC_APROBACION_DEUDA = "fec_aprobacion_deuda"
_FEC_APROBACION_DT = "fec_aprobacion_dt"
_FEC_APROBACION_MAESTRA = "fec_aprobacion_maestra"
_FEC_APROBACION_RIA = "fec_aprobacion_ria"
_FEC_PAGO = "fec_pago"
_FEC_PAGO_DT = "fec_pago_dt"
_FEC_RRAF = "fec_rraf"
_FEC_RRAF_DET = "fec_rraf_det"
_MTO_PAGO = "mto_pago"
_NUM_DOC_RIN = "num_doc_rin"
_NUM_RUC = "num_ruc"
_OPENPYXL = "openpyxl"
_ORD = "__ord"
_PAGO_SUELTO = "PAGO SUELTO"
_PERIODO = "periodo"
_PERIODO_D = "periodo_d"
_PER_DOC = "per_doc"
_TIPOCRUCE = "TipoCruce"


def cargar_excel(path):
    """Carga un Excel intentando varios motores para soportar .xls y .xlsx de forma robusta."""

    motores_a_probar = []

    # 1. Deducir motor prioritario por extensión
    if path.lower().endswith(".xls"):
        motores_a_probar.append("xlrd")  # Prioridad para .xls
        motores_a_probar.append(_OPENPYXL)  # Fallback por si acaso es un xlsx renombrado
    elif path.lower().endswith(".xlsx"):
        motores_a_probar.append(_OPENPYXL)  # Prioridad para .xlsx
        motores_a_probar.append("xlrd")  # Fallback por si acaso es un xls renombrado
    else:
        # Sin extensión clara, probamos openpyxl primero (estándar moderno)
        motores_a_probar = [_OPENPYXL, "xlrd"]

    errores = []

    # 2. Intentar con los motores en orden
    for engine in motores_a_probar:
        try:
            df = pd.read_excel(path, engine=engine)
            df.columns = df.columns.str.strip().str.lower()
            return df
        except Exception as e:
            errores.append(f"Engine '{engine}': {e!s}")

    # 3. Último intento: Automático (sin especificar engine, pandas decide)
    try:
        df = pd.read_excel(path)
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        errores.append(f"Auto: {e!s}")

    # Si llegamos aquí, nada funcionó
    errores_str = " | ".join(errores)
    raise ValueError(f"No se pudo leer el archivo Excel. Detalles técnicos: {errores_str}")


def fmt_fecha(x):
    try:
        val = pd.to_datetime(x, dayfirst=True, errors=_COERCE)
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
    if _COD_TRI in deuda.columns:
        deuda[_COD_TRI] = deuda[_COD_TRI].astype(str).str.zfill(6)
    if _COD_TRIBUTO in pagos.columns:
        pagos[_COD_TRIBUTO] = pagos[_COD_TRIBUTO].astype(str).str.zfill(6)
    # Reemplazo cod_tri_aso -> cod_tri
    if "cod_tri_aso" in deuda.columns and _COD_TRI not in deuda.columns:
        deuda.rename(columns={"cod_tri_aso": _COD_TRI}, inplace=True)


def _validar_columnas(deuda, pagos):
    req_deuda = [_NUM_DOC_RIN, _NUM_RUC, _COD_TRI, _PER_DOC]
    req_pagos = ["ruc", "doc_pago", _DOC_QUE_PAGA, _FEC_PAGO, _PERIODO, _COD_TRIBUTO, _MTO_PAGO]
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

    deuda_ria = deuda[deuda[_NUM_DOC_RIN].astype(str) == str(ria_num)]
    if deuda_ria.empty:
        pagos = pagos[pagos[_DOC_QUE_PAGA].astype(str) == str(ria_num)]
        if pagos.empty:
            raise ValueError(f"No se encontraron pagos asociados a la RIA {ria_num}")
        return pagos, fecha_maestra

    ruc_asociado = deuda_ria[_NUM_RUC].iloc[0]
    fecha_maestra = deuda_ria[_FEC_APROBACION_DT].iloc[0]
    pagos[_FEC_PAGO_DT] = pd.to_datetime(pagos[_FEC_PAGO], dayfirst=True, errors=_COERCE)
    pagos = pagos[pagos["ruc"].astype(str) == str(ruc_asociado)]
    if pd.notna(fecha_maestra):
        pagos = pagos[pagos[_FEC_PAGO_DT] >= fecha_maestra]
    if pagos.empty:
        raise ValueError(f"No existen pagos para el RUC asociado a la RIA {ria_num}")
    return pagos, fecha_maestra


def _calcular_monto_acogido(deuda, ria_num):
    if ria_num and "mto_aco_sol" in deuda.columns:
        temp_deuda = deuda[deuda[_NUM_DOC_RIN].astype(str) == str(ria_num)]
        return pd.to_numeric(temp_deuda["mto_aco_sol"], errors=_COERCE).fillna(0).sum()
    return 0.0


def _construir_base(deuda, pagos, fecha_maestra):
    """Une los pagos con las fechas de aprobación de RIA, RRAF y deuda acogida."""
    base = pagos.copy()
    base[_FEC_APROBACION_MAESTRA] = fecha_maestra

    deuda_ria = deuda[[_NUM_DOC_RIN, _FEC_APROBACION_DT]].drop_duplicates()
    base = base.merge(
        deuda_ria.rename(columns={_NUM_DOC_RIN: "ria_doc", _FEC_APROBACION_DT: _FEC_APROBACION_RIA}),
        left_on=_DOC_QUE_PAGA,
        right_on="ria_doc",
        how="left",
    )

    # RRAF cabecera (080300)
    rraf_cab = deuda[[_NUM_DOC_RIN, _FEC_APROBACION_DT]].drop_duplicates()
    base = base.merge(
        rraf_cab.rename(columns={_NUM_DOC_RIN: "rraf_cab", _FEC_APROBACION_DT: _FEC_RRAF}),
        left_on=_DOC_QUE_PAGA,
        right_on="rraf_cab",
        how="left",
    )

    # Deuda acogida detalle (RRAF)
    rraf_det = deuda[deuda[_COD_TRI].astype(str).str.zfill(6) == "080201"][["num_doc", _FEC_APROBACION_DT]]
    base = base.merge(
        rraf_det.rename(columns={"num_doc": "rraf_det", _FEC_APROBACION_DT: _FEC_RRAF_DET}),
        left_on=_DOC_QUE_PAGA,
        right_on="rraf_det",
        how="left",
    )

    # Código-periodo (RIA normal)
    deuda_codper = deuda[[_NUM_RUC, _COD_TRI, _PER_DOC, _FEC_APROBACION_DT]].drop_duplicates()
    base = base.merge(
        deuda_codper.rename(
            columns={
                _NUM_RUC: "ruc_d",
                _COD_TRI: _COD_TRIBUTO_D,
                _PER_DOC: _PERIODO_D,
                _FEC_APROBACION_DT: _FEC_APROBACION_DEUDA,
            }
        ),
        left_on=["ruc", _COD_TRIBUTO, _PERIODO],
        right_on=["ruc_d", _COD_TRIBUTO_D, _PERIODO_D],
        how="left",
    )

    base[_FEC_PAGO_DT] = pd.to_datetime(base[_FEC_PAGO], dayfirst=True, errors=_COERCE)
    base[_FEC_APROBACION_DEUDA] = (
        base[_FEC_APROBACION_DEUDA].fillna(base[_FEC_APROBACION_RIA]).fillna(base[_FEC_APROBACION_MAESTRA])
    )
    return base


def _clasificar_pago(row):
    cod = str(row.get(_COD_TRIBUTO, "")).zfill(6)

    if pd.notna(row.get(_FEC_RRAF)) and cod in ("080300", "052309", "053302"):
        return "PAGO RRAF"
    if pd.notna(row.get(_FEC_RRAF_DET)) and cod == "080201":
        return "PAGO DEUDA ACOGIDA (RRAF)"
    if pd.notna(row.get(_FEC_APROBACION_RIA)) and cod in ("080201", "052106", "053105"):
        return "PAGO FRACCIONAMIENTO"
    if row.get(_COD_TRIBUTO_D) == row.get(_COD_TRIBUTO) and row.get(_PERIODO_D) == row.get(_PERIODO):
        return "PAGO DEUDA ACOGIDA (RIA)"
    return _PAGO_SUELTO


def _aplicar_filtros_correctivos(base):
    cods = base[_COD_TRIBUTO].astype(str).str.zfill(6)
    base = base[~((base[_TIPOCRUCE] == _PAGO_SUELTO) & (~cods.isin(CODIGOS_RRAF)))]
    if base[_TIPOCRUCE].str.contains("RRAF", na=False).any():
        base = base[base[_COD_TRIBUTO].astype(str).str.zfill(6).isin(CODIGOS_RRAF)]
    return base


def _filtro_final_ria(base, ria_num):
    doc = str(ria_num)
    filtro_directos = base[base[_DOC_QUE_PAGA].astype(str) == doc]
    filtro_deuda = base[base[_TIPOCRUCE].str.contains("DEUDA ACOGIDA", na=False)]
    filtro_suelto = base[
        (base[_COD_TRIBUTO].astype(str).str.zfill(6).isin(CODIGOS_RRAF))
        & (base["doc_pago"].astype(str) == base[_DOC_QUE_PAGA].astype(str))
    ]
    return pd.concat([filtro_directos, filtro_deuda, filtro_suelto]).drop_duplicates()


def _formatear_salida(base):
    base["fec_aprobacion"] = (
        base[_FEC_APROBACION_DEUDA]
        .fillna(base[_FEC_RRAF])
        .fillna(base[_FEC_RRAF_DET])
        .fillna(base[_FEC_APROBACION_MAESTRA])
        .map(fmt_fecha)
    )
    base[_FEC_PAGO] = base[_FEC_PAGO_DT].map(fmt_fecha)

    base[_ORD] = pd.to_datetime(base[_FEC_PAGO], dayfirst=True, errors=_COERCE)
    base = base.sort_values(_ORD).drop(columns=_ORD)

    campos = [_TIPOCRUCE, "ruc", _DOC_QUE_PAGA, "fec_aprobacion", _FEC_PAGO, _PERIODO, _COD_TRIBUTO, _MTO_PAGO]
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
        deuda[_FEC_APROBACION_DT] = pd.to_datetime(deuda.iloc[:, 21], dayfirst=True, errors=_COERCE)
    except Exception:
        deuda[_FEC_APROBACION_DT] = pd.NaT

    deuda["es_rraf"] = False  # placeholder sin uso

    pagos, fecha_maestra = _filtrar_pagos_por_ria(deuda, pagos, ria_num)
    monto_acogido = _calcular_monto_acogido(deuda, ria_num)

    base = _construir_base(deuda, pagos, fecha_maestra)
    base[_TIPOCRUCE] = base.apply(_clasificar_pago, axis=1)
    base = _aplicar_filtros_correctivos(base)
    if ria_num:
        base = _filtro_final_ria(base, ria_num)

    return _formatear_salida(base), monto_acogido


def result_summary(df, monto_acogido=0):
    pagos_deu = len(df[df[_TIPOCRUCE].str.contains("DEUDA ACOGIDA", na=False)])
    pagos_fracc = len(df[df[_TIPOCRUCE].str.contains("FRACCIONAMIENTO", na=False)])
    pagos_suelto = len(df[df[_TIPOCRUCE].str.contains("SUELTO", na=False)])
    pagos_rraf = len(df[df[_TIPOCRUCE].str.contains("RRAF", na=False)])
    total = len(df)
    monto_total = pd.to_numeric(df[_MTO_PAGO], errors=_COERCE).fillna(0).sum()

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
    df = df[df[_TIPOCRUCE].isin(["PAGO FRACCIONAMIENTO", "PAGO DEUDA ACOGIDA (RIA)", _PAGO_SUELTO])].reset_index(
        drop=True
    )
    return df, monto


def ejecutar_cruce_rraf(path_deuda, path_pagos, ria_num=""):
    df, monto = ejecutar_cruce(path_deuda, path_pagos, ria_num, validar_tipo="RRAF")
    df = df[df[_TIPOCRUCE].isin(["PAGO RRAF", "PAGO DEUDA ACOGIDA (RRAF)", _PAGO_SUELTO])].reset_index(drop=True)
    return df, monto
