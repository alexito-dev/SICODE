import asyncio
import json
import logging
import os
import pandas as pd
from datetime import datetime

# Logger para SEACE
logger = logging.getLogger("SEACE-Service")

# URL de la API de exportación
BASE_URL = "https://eap.oece.gob.pe/perfilprov-bus/1.0/ficha/{ruc}/contrataciones/exportar"

import subprocess
import contextlib

# Textos repetidos
_ERROR = "error"
_REASON = "reason"


async def fetch_with_powershell(url):
    """
    MODO FANTASMA: Ejecuta PowerShell usando subprocess.run() en un thread separado.
    Esto evita el NotImplementedError de asyncio en Uvicorn.
    """
    ps_command = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"$url='{url}'; "
        "$wc=New-Object System.Net.WebClient; "
        "$wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); "
        "$wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials; "
        "$wc.DownloadString($url)"
    )

    system_root = os.environ.get("SystemRoot", "C:\\Windows")
    ps_exe = os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    if not os.path.exists(ps_exe):
        ps_exe = "powershell"

    def _run_ps():
        return subprocess.run(
            [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=0x08000000,
            timeout=95,
        )

    try:
        proc = await asyncio.to_thread(_run_ps)

        if proc.returncode != 0:
            err_msg = proc.stderr.strip()
            print(f"[DEBUG PS] Error en PS: {err_msg} | RC: {proc.returncode}")
            return {_ERROR: err_msg or f"Error PS (RC={proc.returncode})"}

        output = proc.stdout.strip()
        if not output:
            return None

        return json.loads(output)
    except Exception as e:
        print(f"[DEBUG PS] Excepcion en _run_ps: {repr(e)}")
        # Check specifically for NotImplementedError string repr
        err_msg = str(e).strip() or repr(e)
        return {_ERROR: f"Error Interno PS: {err_msg}"}


async def fetch_contracts_for_ruc(ruc):
    """Obtiene contratos de forma asíncrona y invisible."""
    all_contracts = []
    page_number = 1
    total_pages = 1

    logger.info(f"[*] Procesando RUC: {ruc} (Power-Bridge Invisible)")

    while page_number <= total_pages:
        url = BASE_URL.format(ruc=ruc) + f"?pageNumber={page_number}&pageSize=500"
        data = await fetch_with_powershell(url)

        if data and _ERROR in data:
            raise Exception(data[_ERROR])

        if not data:
            break

        # Extraer info del proveedor
        proveedor_info = data.get("proveedorE02") or {}
        ruc_prov = proveedor_info.get("ruc_proveedor") or ruc
        nom_prov = proveedor_info.get("razon_social_proveedor") or ""

        # Extraer contratos
        contracts = data.get("contratosE01") or []
        for c in contracts:
            c["_ruc_prov"] = ruc_prov
            c["_nom_prov"] = nom_prov

        all_contracts.extend(contracts)

        # Paginación
        search_info = data.get("searchInfo")
        if not search_info:
            break
        total_pages = search_info.get("pageTotal", 1)

        page_number += 1
        if page_number > total_pages:
            break
        await asyncio.sleep(0.1)  # Pausa técnica

    return all_contracts


async def process_rucs(ruc_list, start_date=None):
    """Procesa múltiples RUCs sin congelar el portal."""
    all_valid_contracts = []
    success_rucs = []
    failed_rucs = []

    len(ruc_list)
    for _i, ruc in enumerate(ruc_list):
        ruc = ruc.strip()

        if not ruc.isdigit() or len(ruc) != 11:
            failed_rucs.append({"ruc": ruc, _REASON: "Formato inválido"})
            continue

        try:
            raw_contracts = await fetch_contracts_for_ruc(ruc)
            if not raw_contracts:
                failed_rucs.append({"ruc": ruc, _REASON: "Sin contratos encontrados"})
                continue

            valid_contracts = process_contracts(raw_contracts, start_date)
            if not valid_contracts:
                failed_rucs.append({"ruc": ruc, _REASON: "Sin contratos en el rango de fecha"})
                continue

            all_valid_contracts.extend(valid_contracts)
            success_rucs.append(ruc)
        except Exception as e:
            err_msg = str(e).strip()
            print(f"[DEBUG PS] Error procesando RUC {ruc}: '{err_msg}'")
            failed_rucs.append({"ruc": ruc, _REASON: err_msg or "Error desconocido"})

    return {"results": all_valid_contracts, "success_rucs": success_rucs, "failed_rucs": failed_rucs}


def process_contracts(contracts, start_date=None):
    """Aplica el filtro de fecha y da formato al monto."""
    processed_data = []

    if isinstance(start_date, str) and start_date.strip():
        try:
            START_DATE = datetime.strptime(start_date, "%Y-%m-%d")
        except Exception:
            START_DATE = datetime(2025, 12, 1)
    else:
        START_DATE = datetime(2025, 12, 1)

    for contract in contracts:
        date_str = contract.get("fecha_prevista_de_fin_de_contrato")
        c_date = None
        if date_str:
            try:
                c_date = datetime.strptime(date_str, "%d/%m/%Y")
            except Exception:
                with contextlib.suppress(Exception):
                    c_date = datetime.strptime(date_str.split("T")[0], "%Y-%m-%d")

        if c_date and c_date >= START_DATE:
            monto_raw = contract.get("monto_del_contrato_original")
            monto_f = ""
            if monto_raw is not None:
                try:
                    monto_f = f"S/. {float(monto_raw):,.2f}"
                except Exception:
                    monto_f = str(monto_raw)

            processed_data.append(
                {
                    "RUC": contract.get("_ruc_prov", ""),
                    "PROVEEDOR": contract.get("_nom_prov", ""),
                    "DESCRIPCION": contract.get("descripcion", ""),
                    "ENTIDAD": contract.get("entidad", ""),
                    "FECHA DE FIRMA DE CONTRATO": contract.get("fecha_de_firma_de_contrato", ""),
                    "FECHA PREVISTA DE FIN DE CONTRATO": contract.get("fecha_prevista_de_fin_de_contrato", ""),
                    "MONTO": monto_f,
                    "OBJETO": contract.get("objeto", ""),
                    "ESTADO": contract.get("estado", ""),
                    "MIEMBROS CONSORCIO": _limpiar_miembros(contract.get("miembros_consorcio", "")),
                }
            )

    return processed_data


def _limpiar_miembros(valor: str) -> str:
    if not valor:
        return valor
    partes = [p for p in str(valor).split("|") if p.strip()]
    if len(partes) <= 1:
        return valor
    return "|".join(partes[:-1])


def generate_excel_bytes(data, failed_rucs=None):
    """Prepara los DataFrames para el Excel final."""
    columns_requested = [
        "RUC",
        "PROVEEDOR",
        "DESCRIPCION",
        "ENTIDAD",
        "FECHA DE FIRMA DE CONTRATO",
        "FECHA PREVISTA DE FIN DE CONTRATO",
        "MONTO",
        "OBJETO",
        "ESTADO",
    ]
    all_columns = columns_requested + ["MIEMBROS CONSORCIO"]

    df_datos = pd.DataFrame(data, columns=all_columns) if data else pd.DataFrame(columns=all_columns)
    df_simple = df_datos[columns_requested].copy() if not df_datos.empty else pd.DataFrame(columns=columns_requested)

    df_fallidos = pd.DataFrame(failed_rucs or [], columns=["ruc", _REASON])
    df_fallidos.columns = ["RUC", "Motivo"]

    return df_datos, df_fallidos, df_simple
