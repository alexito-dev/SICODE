"""
MEFProy Service - Conector y automatizador para la plataforma de Proyectos PTE del MEF.
"""
import urllib.request
import ssl
import json
import csv
import io
import re
import os
import base64
import subprocess
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger("MEFProy-Service")

UNIDADES_EJECUTORAS_PRESET = [
    {"id": "818", "nombre": "GOBIERNO REGIONAL JUNIN - SEDE CENTRAL", "departamento": "Junín"},
    {"id": "728", "nombre": "PROVIAS NACIONAL - MTC", "departamento": "Lima / Nacional"},
    {"id": "796", "nombre": "MUNICIPALIDAD METROPOLITANA DE LIMA", "departamento": "Lima"},
    {"id": "803", "nombre": "GOBIERNO REGIONAL AREQUIPA - SEDE CENTRAL", "departamento": "Arequipa"},
    {"id": "805", "nombre": "GOBIERNO REGIONAL CUSCO - SEDE CENTRAL", "departamento": "Cusco"},
    {"id": "814", "nombre": "GOBIERNO REGIONAL PIURA - SEDE CENTRAL", "departamento": "Piura"},
    {"id": "815", "nombre": "GOBIERNO REGIONAL LA LIBERTAD - SEDE CENTRAL", "departamento": "La Libertad"},
    {"id": "1209", "nombre": "AUTORIDAD NACIONAL DE INFRAESTRUCTURA (ANIN)", "departamento": "Nacional"},
    {"id": "570", "nombre": "MINISTERIO DE EDUCACION - SEDE CENTRAL", "departamento": "Nacional"},
    {"id": "573", "nombre": "MINISTERIO DE SALUD - SEDE CENTRAL", "departamento": "Nacional"}
]

MEF_BASE_URL = "https://ofi5.mef.gob.pe/proyectos_pte/forms/UnidadEjecutora.aspx"


def detectar_sector(nombre: str) -> Dict[str, str]:
    """Clasifica temáticamente el proyecto para enriquecer la visualización."""
    n = nombre.upper()
    if any(k in n for k in ["EDUCACI", "COLEGIO", " I.E", "INSTITUT", "ESCUELA", "AULA"]):
        return {"nombre": "Educación", "icono": "fa-graduation-cap", "color": "blue"}
    if any(k in n for k in ["SALUD", "HOSPITAL", "POSTA", "MEDIC", "CENTRO DE SALUD"]):
        return {"nombre": "Salud", "icono": "fa-heart-pulse", "color": "red"}
    if any(k in n for k in ["CARRETERA", "PUENTE", "VIAL", "CAMINO", "PISTA", "TRANSPORTE", "ASFALT", "PAVIMENT"]):
        return {"nombre": "Transporte / Vías", "icono": "fa-road", "color": "amber"}
    if any(k in n for k in ["AGUA", "SANEAMIENTO", "ALCANTARILLADO", "DESAGUE", "RESERVORIO", "POTABLE"]):
        return {"nombre": "Saneamiento / Agua", "icono": "fa-faucet-drip", "color": "cyan"}
    if any(k in n for k in ["RIEGO", "CANAL", "AGRO", "AGRICOLA", "REPRESA", "PRODUCTIV"]):
        return {"nombre": "Riego / Agro", "icono": "fa-seedling", "color": "green"}
    if any(k in n for k in ["ELECTRI", "ENERGIA", "REDES ELECTRICAS"]):
        return {"nombre": "Electrificación", "icono": "fa-bolt", "color": "yellow"}
    if any(k in n for k in ["SEGURIDAD", "COMISARIA", "MUNICIPAL", "PALACIO", "CIVICO", "INSTITUC"]):
        return {"nombre": "Institucional / Seguridad", "icono": "fa-building-columns", "color": "purple"}
    return {"nombre": "Infraestructura General", "icono": "fa-helmet-safety", "color": "slate"}


class MEFProyService:
    def __init__(self):
        self.ssl_ctx = ssl.create_default_context()
        self.ssl_ctx.check_hostname = False
        self.ssl_ctx.verify_mode = ssl.CERT_NONE

    def _get_headers(self, id_ue: str = "818", periodo: str = "2026") -> Dict[str, str]:
        return {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": "https://ofi5.mef.gob.pe",
            "Referer": f"{MEF_BASE_URL}?tipo=2&IdUE={id_ue}&IdUEBase={id_ue}&periodoBase={periodo}"
        }

    def _call_webmethod_powershell(self, method: str, payload: Dict[str, Any], id_ue: str = "818", periodo: str = "2026") -> Dict[str, Any]:
        """Ejecuta la petición usando PowerShell Ghost Mode para saltar proxies y restricciones de red (como en Boletín/SEACE)."""
        url = f"{MEF_BASE_URL}/{method}"
        payload_str = json.dumps(payload)
        payload_b64 = base64.b64encode(payload_str.encode("utf-8")).decode("ascii")
        referer = f"{MEF_BASE_URL}?tipo=2&IdUE={id_ue}&IdUEBase={id_ue}&periodoBase={periodo}"

        ps_lines = [
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8",
            "[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13",
            "[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}",
            "$wc = New-Object System.Net.WebClient",
            "$wc.Proxy = [System.Net.WebRequest]::GetSystemWebProxy()",
            "$wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials",
            "$wc.Headers.Add('Content-Type', 'application/json; charset=utf-8')",
            "$wc.Headers.Add('Accept', 'application/json, text/javascript, */*; q=0.01')",
            "$wc.Headers.Add('X-Requested-With', 'XMLHttpRequest')",
            "$wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')",
            f"$wc.Headers.Add('Referer', '{referer}')",
            f"$body = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('{payload_b64}'))",
            f"$resp = $wc.UploadString('{url}', 'POST', $body)",
            "[Console]::Out.Write($resp)"
        ]
        ps_script = "\n".join(ps_lines)
        encoded_cmd = base64.b64encode(ps_script.encode("utf-16le")).decode("ascii")

        system_root = os.environ.get("SystemRoot", "C:\\Windows")
        ps_exe = os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
        if not os.path.exists(ps_exe):
            ps_exe = "powershell"

        try:
            proc = subprocess.run(
                [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded_cmd],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=25,
                creationflags=0x08000000
            )
            if proc.returncode == 0 and proc.stdout.strip():
                return json.loads(proc.stdout.strip())
            else:
                raise Exception(f"PowerShell WebClient falló: {proc.stderr.strip() or f'RC={proc.returncode}'}")
        except subprocess.TimeoutExpired:
            raise TimeoutError("El servidor del MEF (ofi5.mef.gob.pe) excedió el tiempo máximo de espera (Timeout de 25s en PowerShell).")

    def _call_webmethod_curl(self, method: str, payload: Dict[str, Any], id_ue: str = "818", periodo: str = "2026") -> Dict[str, Any]:
        """Ejecuta la petición usando curl.exe nativo de Windows (HTTP/2 + Schannel) para velocidad óptima y bypass de WAF F5 del MEF."""
        url = f"{MEF_BASE_URL}/{method}"
        payload_str = json.dumps(payload)
        referer = f"{MEF_BASE_URL}?tipo=2&IdUE={id_ue}&IdUEBase={id_ue}&periodoBase={periodo}"

        system_root = os.environ.get("SystemRoot", "C:\\Windows")
        curl_exe = os.path.join(system_root, "System32", "curl.exe")
        if not os.path.exists(curl_exe):
            curl_exe = "curl"

        cmd = [
            curl_exe, "-4", "--connect-timeout", "8", "--max-time", "20", "--ssl-no-revoke", "-s", "-k", "-X", "POST", url,
            "-H", "Content-Type: application/json; charset=utf-8",
            "-H", "Accept: application/json, text/javascript, */*; q=0.01",
            "-H", "X-Requested-With: XMLHttpRequest",
            "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "-H", f"Referer: {referer}",
            "-d", payload_str
        ]


        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=25,
            creationflags=0x08000000
        )
        if proc.returncode == 0 and proc.stdout.strip():
            raw = proc.stdout.strip()
            if raw.startswith("{") and raw.endswith("}"):
                return json.loads(raw)
            else:
                raise Exception(f"Respuesta de curl no es JSON válido: {raw[:150]}")
        else:
            raise Exception(f"curl.exe falló: {proc.stderr.strip() or f'RC={proc.returncode}'}")

    def _call_webmethod(self, method: str, payload: Dict[str, Any], id_ue: str = "818", periodo: str = "2026") -> Dict[str, Any]:
        """Consulta WebMethod con triple motor y reintentos automáticos contra congestionamiento del MEF."""
        import time

        for intento in range(1, 3):
            # 1. Motor primario ultra-rápido: curl.exe nativo (-4 IPv4)
            try:
                return self._call_webmethod_curl(method, payload, id_ue=id_ue, periodo=periodo)
            except Exception as err_curl:
                logger.warning(f"[Intento {intento}] Motor curl nativo falló ({err_curl}). Intentando conexión directa urllib...")

            # 2. Motor secundario: urllib directa
            try:
                url = f"{MEF_BASE_URL}/{method}"
                data_bytes = json.dumps(payload).encode("utf-8")
                headers = self._get_headers(id_ue=id_ue, periodo=periodo)

                req = urllib.request.Request(url, data=data_bytes, headers=headers)
                with urllib.request.urlopen(req, context=self.ssl_ctx, timeout=12) as resp:
                    raw = resp.read().decode("utf-8-sig", errors="ignore")
                    return json.loads(raw)
            except Exception as err_urllib:
                logger.warning(f"[Intento {intento}] Conexión urllib falló ({err_urllib}). Conmutando a modo PowerShell...")

            # 3. Motor terciario: PowerShell Ghost Mode
            try:
                return self._call_webmethod_powershell(method, payload, id_ue=id_ue, periodo=periodo)
            except Exception as err_ps:
                logger.warning(f"[Intento {intento}] Motor PowerShell falló: {err_ps}")

            time.sleep(1)

        raise TimeoutError("La plataforma pública del MEF (ofi5.mef.gob.pe) se encuentra temporalmente lenta o saturada. Por favor vuelve a hacer clic en 'Consultar'.")


    def obtener_presets_ue(self) -> List[Dict[str, str]]:
        return UNIDADES_EJECUTORAS_PRESET

    def obtener_periodos(self, p_periodo: str = "Ejecucion") -> List[str]:
        try:
            res = self._call_webmethod("GetPeriodo", {"pPeriodo": p_periodo})
            items = res.get("d", [])
            periodos = [item.get("Valor") for item in items if item.get("Valor")]
            periodos.sort(reverse=True)
            return periodos if periodos else ["2026", "2025", "2024", "2023", "2022", "2021", "2020"]
        except Exception:
            return ["2026", "2025", "2024", "2023", "2022", "2021", "2020"]

    def obtener_ejecucion(
        self,
        id_ue: str = "818",
        periodo: str = "2026",
        page_size: int = 50,
        page_number: int = 1,
        sort_column: str = "codDGPP",
        sort_order: str = "asc",
        nom_proyecto: str = "",
        cod_dgpp: str = "",
        cod_snip: str = ""
    ) -> Dict[str, Any]:
        """Consulta la ejecución presupuestal de proyectos en el MEF."""
        payload = {
            "pPageSize": page_size,
            "pPageNumber": page_number,
            "pSortColumn": sort_column,
            "pSortOrder": sort_order,
            "pPeriodo": periodo,
            "pCodEjecutora": id_ue,
            "pCodDGPP": cod_dgpp,
            "pCodSNIP": cod_snip,
            "pNomProyecto": nom_proyecto
        }

        res = self._call_webmethod("GetEjecucion", payload, id_ue=id_ue, periodo=periodo)
        data_grid = res.get("d", {})
        
        raw_rows = data_grid.get("rows", [])
        total_records = data_grid.get("records", 0)
        total_pages = data_grid.get("total", 0)
        current_page = data_grid.get("page", page_number)

        parsed_rows = []
        total_costo_p = 0.0
        total_pim_p = 0.0
        total_devengado_p = 0.0

        for row in raw_rows:
            cells = row.get("cell", [])
            if len(cells) >= 8:
                try:
                    c_dgpp = str(cells[0]).strip()
                    c_snip = str(cells[1]).strip()
                    nombre = str(cells[2]).strip()
                    
                    costo_proy = float(cells[3]) if cells[3] else 0.0
                    pim = float(cells[4]) if cells[4] else 0.0
                    devengado = float(cells[5]) if cells[5] else 0.0
                    
                    avance_pct = (devengado / pim * 100.0) if pim > 0 else 0.0
                    dev_acumulado = float(cells[7]) if len(cells) > 7 and cells[7] else 0.0

                    total_costo_p += costo_proy
                    total_pim_p += pim
                    total_devengado_p += devengado

                    sector = detectar_sector(nombre)
                    cod_ref = c_dgpp or c_snip
                    ficha_url = f"https://ofi5.mef.gob.pe/invierteweb/Consulta/ConsultaFicha/ConsultaFicha?CodigoUnico={cod_ref}" if cod_ref else None

                    parsed_rows.append({
                        "id": row.get("id"),
                        "cod_dgpp": c_dgpp,
                        "cod_snip": c_snip,
                        "nom_proyecto": nombre,
                        "costo_proyecto": costo_proy,
                        "pim": pim,
                        "devengado": devengado,
                        "avance_porcentaje": round(avance_pct, 2),
                        "devengado_acumulado": dev_acumulado,
                        "sector": sector,
                        "ficha_url": ficha_url
                    })
                except Exception:
                    continue

        avance_global = (total_devengado_p / total_pim_p * 100.0) if total_pim_p > 0 else 0.0

        return {
            "status": "success",
            "id_ue": id_ue,
            "periodo": periodo,
            "page": current_page,
            "total_pages": total_pages,
            "total_records": total_records,
            "page_records": len(parsed_rows),
            "totales_pagina": {
                "pim_total": round(total_pim_p, 2),
                "devengado_total": round(total_devengado_p, 2),
                "avance_global": round(avance_global, 2),
                "costo_proyectos_total": round(total_costo_p, 2)
            },
            "proyectos": parsed_rows
        }

    def obtener_banco_proyectos(
        self,
        id_ue: str = "818",
        periodo: str = "2026",
        page_size: int = 50,
        page_number: int = 1,
        sort_column: str = "codDGPP",
        sort_order: str = "asc"
    ) -> Dict[str, Any]:
        """Consulta el Banco de Proyectos (DGPI) con fechas de registro y viabilidad."""
        payload = {
            "pPageSize": page_size,
            "pPageNumber": page_number,
            "pSortColumn": sort_column,
            "pSortOrder": sort_order,
            "pPeriodo": periodo,
            "pCodEjecutora": id_ue
        }

        res = self._call_webmethod("GetProyectos", payload, id_ue=id_ue, periodo=periodo)
        data_grid = res.get("d", {})
        raw_rows = data_grid.get("rows", [])
        total_records = data_grid.get("records", 0)
        total_pages = data_grid.get("total", 0)

        proyectos = []
        for row in raw_rows:
            cells = row.get("cell", [])
            if len(cells) >= 7:
                try:
                    c_dgpp = str(cells[0]).strip()
                    c_snip = str(cells[1]).strip()
                    nombre = str(cells[2]).strip()
                    f_registro = str(cells[3]).strip()
                    f_viabilidad = str(cells[4]).strip()
                    monto_viab = float(cells[5]) if cells[5] else 0.0
                    costo_p = float(cells[6]) if cells[6] else 0.0

                    sector = detectar_sector(nombre)
                    cod_ref = c_dgpp or c_snip
                    ficha_url = f"https://ofi5.mef.gob.pe/invierteweb/Consulta/ConsultaFicha/ConsultaFicha?CodigoUnico={cod_ref}" if cod_ref else None

                    proyectos.append({
                        "id": row.get("id"),
                        "cod_dgpp": c_dgpp,
                        "cod_snip": c_snip,
                        "nom_proyecto": nombre,
                        "fecha_registro": f_registro,
                        "fecha_viabilidad": f_viabilidad,
                        "monto_viabilidad": monto_viab,
                        "costo_proyecto": costo_p,
                        "sector": sector,
                        "ficha_url": ficha_url
                    })
                except Exception:
                    continue

        return {
            "status": "success",
            "id_ue": id_ue,
            "periodo": periodo,
            "total_records": total_records,
            "total_pages": total_pages,
            "proyectos": proyectos
        }

    def obtener_ranking(self, id_ue: str = "818", periodo: str = "2026") -> List[Dict[str, Any]]:
        try:
            res = self._call_webmethod("GetRankingProyecto", {"pPeriodo": periodo, "pCodEjecutora": id_ue}, id_ue=id_ue, periodo=periodo)
            items = res.get("d", [])
            ranking = []
            for it in items:
                nombre = it.get("nombre", "")
                ranking.append({
                    "cod_dgpp": it.get("dgpp", ""),
                    "cod_snip": it.get("snip", ""),
                    "nombre": nombre,
                    "pim": float(it.get("pim", 0)),
                    "sector": detectar_sector(nombre)
                })
            return ranking
        except Exception:
            return []

    def obtener_detalle_proyecto(self, id_ue: str, cod_proyecto: str) -> List[Dict[str, Any]]:
        payload = {
            "pPageSize": 50,
            "pPageNumber": 1,
            "pSortColumn": "periodo",
            "pSortOrder": "desc",
            "pCodEjecutora": id_ue,
            "pCodProyecto": cod_proyecto
        }

        res = self._call_webmethod("GetDetalle", payload, id_ue=id_ue)
        data_grid = res.get("d", {})
        raw_rows = data_grid.get("rows", [])

        detalles = []
        for r in raw_rows:
            cells = r.get("cell", [])
            if len(cells) >= 8:
                try:
                    detalles.append({
                        "periodo": str(cells[0]),
                        "pim": float(cells[1]) if cells[1] else 0.0,
                        "trim1": float(cells[2]) if cells[2] else 0.0,
                        "trim2": float(cells[3]) if cells[3] else 0.0,
                        "trim3": float(cells[4]) if cells[4] else 0.0,
                        "trim4": float(cells[5]) if cells[5] else 0.0,
                        "total_devengado": float(cells[6]) if cells[6] else 0.0,
                        "avance": float(cells[7]) if cells[7] else 0.0
                    })
                except Exception:
                    continue

        return detalles

    def generar_excel_xlsx(self, proyectos: List[Dict[str, Any]], id_ue: str, periodo: str, tipo_vista: str = "ejecucion") -> bytes:
        """Genera un archivo XLSX ejecutivo de alto impacto visual con cabecera agrupada, fórmulas y formato financiero."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Proyectos MEF {periodo}"
        ws.views.sheetView[0].showGridLines = True

        # Buscar datos institucionales
        ue_info = next((u for u in UNIDADES_EJECUTORAS_PRESET if u["id"] == id_ue), None)
        inst_nombre = ue_info["nombre"] if ue_info else f"UNIDAD EJECUTORA N° {id_ue}"
        dep_nombre = ue_info["departamento"] if ue_info else ""

        # Paleta de Colores Armónica Exacta a la Grilla
        c_hdr_bg = "E4F1FC"        # Azul pastel suave
        c_hdr_group_bg = "D4ECFC"  # Azul pastel para grupo año
        c_hdr_text = "033F6C"      # Azul texto institucional
        c_hdr_border = "BEDBFF"    # Borde azul claro
        c_total_bg = "EBF4FC"      # Fondo totales
        c_zebra = "F8FAFC"         # Fila cebra

        header_font = Font(name="Calibri", size=10, bold=True, color=c_hdr_text)
        sub_header_font = Font(name="Calibri", size=9.5, bold=True, color=c_hdr_text)
        group_font = Font(name="Calibri", size=10.5, bold=True, color=c_hdr_text)
        title_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        subtitle_font = Font(name="Calibri", size=11, bold=True, color="014F88")
        meta_font = Font(name="Calibri", size=9, bold=False, color="475569")
        bold_font = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
        regular_font = Font(name="Calibri", size=9.5, color="1E293B")
        code_font = Font(name="Consolas", size=9.5, bold=True, color="0163AC")

        header_border = Border(
            left=Side(style='thin', color=c_hdr_border),
            right=Side(style='thin', color=c_hdr_border),
            top=Side(style='thin', color=c_hdr_border),
            bottom=Side(style='thin', color=c_hdr_border)
        )
        thin_border = Border(
            left=Side(style='thin', color="E2E8F0"),
            right=Side(style='thin', color="E2E8F0"),
            top=Side(style='thin', color="E2E8F0"),
            bottom=Side(style='thin', color="E2E8F0")
        )
        total_top_border = Border(
            left=Side(style='thin', color=c_hdr_border),
            right=Side(style='thin', color=c_hdr_border),
            top=Side(style='medium', color="0163AC"),
            bottom=Side(style='double', color="0163AC")
        )

        # 1. ENCABEZADO INSTITUCIONAL
        num_cols = 8 if tipo_vista == "ejecucion" else 7
        last_col_letter = get_column_letter(num_cols)

        # Banner Superior
        ws.merge_cells(f"A1:{last_col_letter}1")
        banner_cell = ws["A1"]
        banner_cell.value = "SISTEMA DE CONSULTA DE INVERSIÓN PÚBLICA (MEF - SIAF / PTE)"
        banner_cell.font = title_font
        banner_cell.fill = PatternFill(start_color="0163AC", end_color="0163AC", fill_type="solid")
        banner_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        # Institución / Gobierno Regional
        ws.merge_cells(f"A2:{last_col_letter}2")
        inst_cell = ws["A2"]
        inst_cell.value = f"GOBIERNO REGIONAL / ENTIDAD: {inst_nombre}" + (f" ({dep_nombre.upper()})" if dep_nombre else "")
        inst_cell.font = subtitle_font
        inst_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Subtítulo y Metadatos
        ws.merge_cells(f"A3:{last_col_letter}3")
        meta_cell = ws["A3"]
        meta_cell.value = f"UNIDAD EJECUTORA: {id_ue} | AÑO FISCAL: {periodo} | VISTA: {tipo_vista.upper()} | FECHA DE EMISIÓN: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        meta_cell.font = meta_font
        meta_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 18

        # 2. CABECERA AGRUPADA DE 2 FILAS (FILA 5 Y 6)
        r5 = 5
        r6 = 6

        if tipo_vista == "ejecucion":
            headers_r5 = {
                "A": "CÓDIGO\nÚNICO",
                "B": "CÓDIGO\nSNIP",
                "C": "PROYECTO",
                "D": "COSTO DEL\nPROYECTO",
                "E": f"{periodo}",
                "H": f"DEVENGADO\nACUMULADO\nAL {periodo}"
            }
            headers_r6 = {
                "E": "PIM",
                "F": "DEVENGADO",
                "G": "AVANCE\n(%)"
            }

            # Merges verticales A5:A6, B5:B6, C5:C6, D5:D6, H5:H6
            for col_l in ["A", "B", "C", "D", "H"]:
                ws.merge_cells(f"{col_l}{r5}:{col_l}{r6}")
                cell = ws[f"{col_l}{r5}"]
                cell.value = headers_r5[col_l]
                cell.fill = PatternFill(start_color=c_hdr_bg, fill_type="solid")
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Merge horizontal E5:G5 (Año)
            ws.merge_cells(f"E{r5}:G{r5}")
            cell_group = ws[f"E{r5}"]
            cell_group.value = headers_r5["E"]
            cell_group.fill = PatternFill(start_color=c_hdr_group_bg, fill_type="solid")
            cell_group.font = group_font
            cell_group.alignment = Alignment(horizontal="center", vertical="center")

            # Subcolumnas E6, F6, G6
            for col_l, title in headers_r6.items():
                cell = ws[f"{col_l}{r6}"]
                cell.value = title
                cell.fill = PatternFill(start_color=c_hdr_bg, fill_type="solid")
                cell.font = sub_header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Bordes para todas las celdas de cabecera
            for r in [r5, r6]:
                for c in range(1, 9):
                    ws.cell(row=r, column=c).border = header_border

        else: # Banco de Proyectos
            headers_banco = [
                "CÓDIGO ÚNICO", "CÓDIGO SNIP", "PROYECTO",
                "FECHA REGISTRO", "FECHA VIABILIDAD", "MONTO VIABILIDAD", "COSTO TOTAL"
            ]
            for col_idx, h in enumerate(headers_banco, 1):
                col_l = get_column_letter(col_idx)
                ws.merge_cells(f"{col_l}{r5}:{col_l}{r6}")
                cell = ws[f"{col_l}{r5}"]
                cell.value = h
                cell.fill = PatternFill(start_color=c_hdr_bg, fill_type="solid")
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                for r in [r5, r6]:
                    ws.cell(row=r, column=col_idx).border = header_border

        ws.row_dimensions[r5].height = 22
        ws.row_dimensions[r6].height = 22

        # 3. FILAS DE DATOS
        curr_row = 7
        for idx, p in enumerate(proyectos):
            is_even = (idx % 2 == 1)
            row_fill = PatternFill(start_color=c_zebra if is_even else "FFFFFF", fill_type="solid")

            if tipo_vista == "banco":
                row_data = [
                    p.get("cod_dgpp", "") or p.get("cod_snip", ""),
                    p.get("cod_snip", ""),
                    p.get("nom_proyecto", ""),
                    p.get("fecha_registro", ""),
                    p.get("fecha_viabilidad", ""),
                    float(p.get("monto_viabilidad", 0) or 0),
                    float(p.get("costo_proyecto", 0) or 0)
                ]
            else:
                pim_val = float(p.get("pim", 0) or 0)
                dev_val = float(p.get("devengado", 0) or 0)
                av_val = (dev_val / pim_val) if pim_val > 0 else 0.0
                row_data = [
                    p.get("cod_dgpp", "") or p.get("cod_snip", ""),
                    p.get("cod_snip", ""),
                    p.get("nom_proyecto", ""),
                    float(p.get("costo_proyecto", 0) or 0),
                    pim_val,
                    dev_val,
                    av_val,
                    float(p.get("devengado_acumulado", 0) or 0)
                ]

            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=col_idx, value=val)
                cell.fill = row_fill
                cell.border = thin_border

                if col_idx == 1:
                    cell.font = code_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 2:
                    cell.font = regular_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 3:
                    cell.font = regular_font
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    cell.font = regular_font
                    if tipo_vista == "ejecucion":
                        if col_idx in [4, 5, 6, 8]:
                            cell.number_format = '"S/" #,##0.00'
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                        elif col_idx == 7:
                            cell.number_format = '0.0%'
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        if col_idx in [6, 7]:
                            cell.number_format = '"S/" #,##0.00'
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                        elif col_idx in [4, 5]:
                            cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[curr_row].height = 20
            curr_row += 1

        # 4. FILA DE TOTALES
        if tipo_vista == "ejecucion" and curr_row > 7:
            ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=3)
            tot_cell = ws.cell(row=curr_row, column=1, value="TOTAL GENERAL:")
            tot_cell.font = bold_font
            tot_cell.alignment = Alignment(horizontal="right", vertical="center")

            for c in range(1, 4):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = total_top_border
                cell.fill = PatternFill(start_color=c_total_bg, fill_type="solid")

            for col_idx in [4, 5, 6, 8]:
                col_letter = get_column_letter(col_idx)
                cell = ws.cell(row=curr_row, column=col_idx)
                cell.value = f"=SUM({col_letter}7:{col_letter}{curr_row-1})"
                cell.font = bold_font
                cell.number_format = '"S/" #,##0.00'
                cell.border = total_top_border
                cell.fill = PatternFill(start_color=c_total_bg, fill_type="solid")
                cell.alignment = Alignment(horizontal="right", vertical="center")

            cell_av = ws.cell(row=curr_row, column=7)
            cell_av.value = f"=IF(E{curr_row}>0, F{curr_row}/E{curr_row}, 0)"
            cell_av.font = bold_font
            cell_av.number_format = '0.0%'
            cell_av.border = total_top_border
            cell_av.fill = PatternFill(start_color=c_total_bg, fill_type="solid")
            cell_av.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[curr_row].height = 24

        # 5. ANCHO ÓPTIMO DE COLUMNAS
        col_widths = {
            "A": 16,  # CÓDIGO ÚNICO
            "B": 15,  # CÓDIGO SNIP
            "C": 55,  # PROYECTO DE INVERSIÓN (espacio amplio)
            "D": 22,  # COSTO DEL PROYECTO
            "E": 20,  # PIM
            "F": 20,  # DEVENGADO
            "G": 15,  # AVANCE %
            "H": 26   # DEVENGADO ACUMULADO
        }
        for col_l, width in col_widths.items():
            if ord(col_l) - ord('A') + 1 <= num_cols:
                ws.column_dimensions[col_l].width = width

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

