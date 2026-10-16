import os
import re
import uuid
import time
import shutil
import zipfile
import json
import urllib.request
from decimal import Decimal
from typing import List
from fastapi import APIRouter, Request, Depends, UploadFile, File, Form
from fastapi.responses import JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
import pdfplumber
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from app.modulos.auth.dependencies import login_required
from app.core.config import settings, sicode_dir

router = APIRouter(
    prefix="/ribot",
    tags=["ribot"],
    dependencies=[Depends(login_required)]
)

templates_dir = settings.get_template_path("app/modulos/ribot/templates")
templates_base = settings.get_template_path("app/modulos/inicio/templates")
templates = Jinja2Templates(directory=[templates_dir, templates_base])

# Directorio temporal de RIbot
TEMP_DIR = sicode_dir("DATA", "RIBOT", "TEMP")

# Directorio central de documentos locales (Estilo Bpsearch)
RIBOT_BASE_DIR = sicode_dir("DATA", "RIBOT", "DOCS")

# Directorio de Plantillas Word de RIbot
def obtener_directorio_plantillas() -> str:
    candidato = r"D:\SICODE\APPS\RIBOT"
    if os.path.exists(candidato):
        return candidato
    local = os.path.join(os.path.dirname(__file__), "plantillas")
    os.makedirs(local, exist_ok=True)
    return local

RIBOT_PLANTILLAS_DIR = obtener_directorio_plantillas()

# ── Sanitización y Limpieza de Archivos ──────────────────────────────────────
def sanitizar_nombre_archivo(nombre: str) -> str:
    r"""Remueve caracteres inválidos en Windows (\ / : * ? " < > |) para evitar crashes OSError."""
    if not nombre:
        return "DOCUMENTO"
    limpio = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', str(nombre))
    limpio = re.sub(r'_+', '_', limpio).strip(' ._')
    return limpio if limpio else "DOCUMENTO"

def limpiar_archivos_temporales(max_horas=6):
    """Elimina carpetas temporales antiguas en TEMP_DIR para evitar fugas de almacenamiento."""
    try:
        ahora = time.time()
        limite = ahora - (max_horas * 3600)
        for item in os.listdir(TEMP_DIR):
            item_path = os.path.join(TEMP_DIR, item)
            if os.path.isdir(item_path):
                st_time = os.path.getmtime(item_path)
                if st_time < limite:
                    shutil.rmtree(item_path, ignore_errors=True)
    except Exception as e:
        print(f"Error en rutina de limpieza de temporales RIbot: {e}")

# ── Configuración de Fechas SUNAT ──────────────────────────────────────────────
MESES = {
    "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
    "05": "mayo",  "06": "junio",   "07": "julio", "08": "agosto",
    "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre"
}

def fecha_larga(fecha_str):
    """Convierte '09/03/2026' → '09 de marzo de 2026' de forma segura."""
    if not fecha_str or not isinstance(fecha_str, str):
        return "[COMPLETAR]"
    try:
        partes = fecha_str.strip().split("/")
        if len(partes) == 3:
            d, m, a = partes
            mes_nombre = MESES.get(m.zfill(2), m)
            return f"{d} de {mes_nombre} de {a}"
    except Exception:
        pass
    return fecha_str

# ── Auxiliares para Generación de Word (python-docx) ──────────────────────────
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_cell_borders(cell):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)

def agregar_tabla(doc, encabezados, fila_datos, col_widths_cm):
    table = doc.add_table(rows=2, cols=len(encabezados))
    table.style = "Table Grid"

    # Fila encabezado
    for i, texto in enumerate(encabezados):
        cell = table.rows[0].cells[i]
        cell.width = Cm(col_widths_cm[i])
        set_cell_bg(cell, "D9D9D9")
        set_cell_borders(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(texto)
        run.bold = True
        run.font.size = Pt(9)

    # Fila datos
    for i, texto in enumerate(fila_datos):
        cell = table.rows[1].cells[i]
        cell.width = Cm(col_widths_cm[i])
        set_cell_borders(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(texto))
        if i == 0:
            run.bold = True
        run.font.size = Pt(9)

    return table

def parrafo_justificado(doc, partes, sangria_cm=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if sangria_cm:
        p.paragraph_format.first_line_indent = Cm(sangria_cm)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.space_before = Pt(0)
    for texto, bold in partes:
        run = p.add_run(texto)
        run.bold = bold
        run.font.size = Pt(10)
    return p

NOMBRES_TRIBUTO = {
    "053100": "SNP Ley 19990",
}

# ── Extractor Ficha del Valor PDF ───────────────────────────────────────────
def extraer_datos_pdf(ruta_pdf):
    datos = {campo: "[COMPLETAR]" for campo in [
        "archivo", "ruc", "razon_social", "domicilio",
        "nro_resolucion", "ciudad", "fecha_resolucion",
        "nro_expediente", "fecha_expediente",
        "nro_orden_pago", "periodo", "codigo_tributo", "nombre_tributo",
        "tributo_original", "interes_original", "total_original",
        "fecha_emision_op", "fecha_declaracion", "nro_orden_declaracion",
        "fecha_rectificatoria", "nro_orden_rectificatoria",
        "tributo_modificado", "interes_modificado", "total_modificado",
        "fecha_interes", "nro_formulario_pago", "nro_boleta_pago",
        "pago", "fecha_pago", "intendencia",
        "codigo_tributo_asoc", "multa", "interes_multa", "total_multa",
        "multa_mod", "interes_multa_mod", "total_multa_mod"
    ]}
    datos["avisos_extraccion"] = []
    datos["archivo"] = os.path.basename(ruta_pdf)

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e:
        print(f"Error parseando PDF en RIbot: {e}")
        return datos

    if not texto:
        return datos

    # 1. Intendencia y Ciudad
    m = re.search(r"^(I\.T\.I\.[^\n]+)", texto, re.MULTILINE)
    if m:
        datos["intendencia"] = m.group(1).split("Ficha del Valor")[0].strip()
    else:
        m_alt = re.search(r"^(.+?)\s+Ficha del Valor", texto, re.MULTILINE)
        if m_alt:
            datos["intendencia"] = m_alt.group(1).strip()
            
    intend = datos.get("intendencia", "").upper()
    if "JUNIN" in intend or "HUANCAYO" in intend:
        datos["ciudad"] = "Huancayo"
    elif "LIMA" in intend:
        datos["ciudad"] = "Lima"
    else:
        datos["ciudad"] = "Huancayo"

    # 2. RUC y Razón Social (En Ficha del Valor viene en la misma línea: RUC 20602398871 HICOGROUP S.A.C.)
    m_ruc = re.search(r"RUC\s*[:\s]*\b((?:10|15|16|17|20)\d{9})\b\s*([^\n\r]+)", texto, re.IGNORECASE)
    if m_ruc:
        datos["ruc"] = m_ruc.group(1).strip()
        rs_candidata = m_ruc.group(2).strip()
        rs_candidata = re.split(r"(?:Notificaci|Montos|Periodo|REC|Estado)", rs_candidata, flags=re.IGNORECASE)[0].strip()
        # La razón social puede continuar en la línea siguiente (ej. "... SR. DE" / "ATACO SOCIEDAD ANONIMA CERRADA")
        lineas = texto[m_ruc.end():].lstrip("\r\n").splitlines()
        if lineas and re.fullmatch(r"[A-ZÁÉÍÓÚÑ0-9&.,\-()/ ]{3,}", lineas[0].strip()) \
                and not re.match(r"(?:Notificaci|Montos|Periodo|REC|Estado|Datos|C[aá]lculo|Fecha|Tributo|Emisi|Etapa)", lineas[0].strip(), re.IGNORECASE):
            rs_candidata += " " + lineas[0].strip()
        datos["razon_social"] = re.sub(r"\s+", " ", rs_candidata)
    else:
        m_ruc2 = re.search(r"\b((?:10|15|16|17|20)\d{9})\b", texto)
        if m_ruc2:
            datos["ruc"] = m_ruc2.group(1).strip()
            for linea in texto.splitlines():
                if datos["ruc"] in linea:
                    partes = linea.split(datos["ruc"])
                    if len(partes) > 1 and partes[1].strip():
                        datos["razon_social"] = re.sub(r"\s+", " ", partes[1].strip())
                        break

    # 3. N° Orden de Pago (13 dígitos)
    m_op = re.search(r"Valor\s+(?:Nro\.?\s*)?(\d{13})", texto, re.IGNORECASE)
    if not m_op:
        m_op = re.search(r"Valor\s+(\d{13})\s+Orden\s+de\s+Pago", texto, re.IGNORECASE)
    if not m_op:
        m_op = re.search(r"Valor Nro\.\s+(\d+)", texto)
    if m_op:
        datos["nro_orden_pago"] = m_op.group(1).strip()

    # 4. Periodo
    m_per = re.search(r"Periodo\s+(\d{4}-\d{2})", texto, re.IGNORECASE)
    if m_per:
        datos["periodo"] = m_per.group(1).strip()

    # 5. Fecha emisión OP
    m_em = re.search(r"Emisi[oó]n\s+(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if m_em:
        datos["fecha_emision_op"] = m_em.group(1).strip()
        datos["fecha_declaracion"] = m_em.group(1).strip()
        datos["fecha_interes"] = m_em.group(1).strip()

    m_fint = re.search(r"Inicio\s+Inter[eé]s\s+(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if m_fint:
        datos["fecha_interes"] = m_fint.group(1).strip()

    # Tributo Asociado (solo en Resoluciones de Multa): "Tributo Asociado 053100 PENREG"
    m_asoc = re.search(r"Tributo\s+Asociado\s+(\d{6})", texto, re.IGNORECASE)
    if m_asoc:
        datos["codigo_tributo_asoc"] = m_asoc.group(1).strip()

    # 6. Código de Tributo
    m_cod = re.search(r"Tributo\s+(\d{6})", texto, re.IGNORECASE)
    if not m_cod:
        m_cod = re.search(r"\b(\d{6})\s+[A-Z0-9]{3,}", texto)
    if m_cod:
        cod = m_cod.group(1).strip()
        datos["codigo_tributo"] = cod
        datos["nombre_tributo"] = {
            "010101": "Impuesto General a las Ventas",
            "011001": "Impuesto a la Renta de Tercera Categoría",
            "053100": "SNP Ley 19990",
        }.get(cod, f"Tributo {cod}")
    else:
        datos["codigo_tributo"] = "010101"
        datos["nombre_tributo"] = "Impuesto General a las Ventas"

    # 7. Montos Originales
    m_montos = re.search(r"C[aá]lculo\s+Saldo\s+Total.*?\nFecha\s+(\d{2}/\d{2}/\d{4}).*?\nTributo\s+([\d,]+\.\d{2}).*?\n.*?Inter[eé]s\s+([\d,]+\.\d{2}).*?\nTotal\s+([\d,]+\.\d{2})", texto, re.IGNORECASE | re.DOTALL)
    if m_montos:
        datos["tributo_original"] = m_montos.group(2).strip()
        datos["interes_original"] = m_montos.group(3).strip()
        datos["total_original"]   = m_montos.group(4).strip()
    else:
        m_op_ev = re.search(r"Emisi[oó]n\s+de\s+OP\s+\d+\s+([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", texto, re.IGNORECASE)
        if m_op_ev:
            datos["tributo_original"] = m_op_ev.group(1).strip()
            datos["interes_original"] = m_op_ev.group(2).strip()
            datos["total_original"]   = m_op_ev.group(3).strip()

    # 8. Declaración Inicial (Formulario y Documento)
    m_dec = re.search(r"Formulario\s+(\d+)\s+Documento\s+(\d+)", texto, re.IGNORECASE)
    if m_dec:
        datos["nro_orden_declaracion"] = m_dec.group(2).strip()
    else:
        m_dec2 = re.search(r"Formulario\s+621\s+(\d+)", texto, re.IGNORECASE)
        if m_dec2:
            datos["nro_orden_declaracion"] = m_dec2.group(1).strip()

    # 9. Etapas: Exp.Error Material (Inicio = Expediente, Fin = Resolución)
    m_et = re.search(r"Exp\.?\s*Error\s+Material\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+\d+\s+(\d+)\s+(\d+)", texto, re.IGNORECASE)
    if m_et:
        datos["fecha_expediente"] = m_et.group(1).strip()
        datos["fecha_resolucion"] = fecha_larga(m_et.group(2).strip())
        datos["nro_expediente"]   = m_et.group(3).strip()
        datos["nro_resolucion"]   = m_et.group(4).strip()
    else:
        m_exp_ev = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}\s+Emisi[oó]n\s+de\s+Expediente\s+error\s+material\s+\d+\s+(\d+)", texto, re.IGNORECASE)
        if m_exp_ev:
            datos["fecha_expediente"] = m_exp_ev.group(1).strip()
            datos["nro_expediente"]   = m_exp_ev.group(2).strip()

    # 10. Montos Modificados (Evento Emisión de R. Error Material c Expediente)
    m_rem = re.search(r"Emisi[oó]n\s+de\s+R\.?\s*Error\s+Material\s+c\s+Expediente\s+\d+\s+(?:\d+\s+)?([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", texto, re.IGNORECASE)
    if m_rem:
        datos["tributo_modificado"] = m_rem.group(1).strip()
        datos["interes_modificado"] = m_rem.group(2).strip()
        datos["total_modificado"]   = m_rem.group(3).strip()
    else:
        m_mov = re.search(r"Ajuste\s+por\s+R\.?\s*Error\s+Material\s+c\s+Expe.*?\(?([\d,]+\.\d{2})\)?\s+[\d,]+\.\d{2}\s+\(?([\d,]+\.\d{2})\)?\s+\(?([\d,]+\.\d{2})\)?", texto, re.IGNORECASE)
        if m_mov:
            datos["tributo_modificado"] = m_mov.group(1).strip()
            datos["interes_modificado"] = m_mov.group(2).strip()
            datos["total_modificado"]   = m_mov.group(3).strip()

    # 10b. Movimientos de Saldo: multa, interés y total (original y modificado)
    #  - Fila "Deuda por RM"  -> valores originales de la multa.
    #  - Filas hasta el último "Ajuste por R. Error Material" -> acumulado = valores modificados.
    #  La celda Tributo puede quedar pisada por la descripción en el PDF, por eso se leen solo los
    #  últimos 4 montos (Capitalizado, Interés, Total, Saldo) y Tributo = Total - Capitalizado - Interés.
    def _dec(s):
        s = s.replace(",", "")
        neg = s.startswith("(")
        v = Decimal(s.strip("()"))
        return -v if neg else v

    def _fmt(v):
        return f"{v:,.2f}"

    seccion = re.split(r"Movimientos\s+de\s+Saldo", texto, flags=re.IGNORECASE)[-1] if re.search(r"Movimientos\s+de\s+Saldo", texto, re.IGNORECASE) else ""
    movs = []
    for linea in seccion.splitlines():
        m_fila = re.match(r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+(.*)$", linea.strip())
        if not m_fila:
            continue
        cuerpo = re.sub(r"\s+\d+\s+\d{13}\s+\d{2}/\d{2}/\d{4}.*$", "", m_fila.group(1))
        montos = re.findall(r"\(?[\d,]+\.\d{2}\)?", cuerpo)
        if len(montos) < 4:
            continue
        cap, inte, tot, sal = (_dec(x) for x in montos[-4:])
        movs.append({"desc": cuerpo, "trib": tot - cap - inte, "int": inte, "tot": tot, "sal": sal})

    if movs:
        orig = next((m for m in movs if re.match(r"Deuda", m["desc"], re.IGNORECASE)), movs[0])
        datos["multa"] = _fmt(orig["trib"])
        datos["interes_multa"] = _fmt(orig["int"])
        datos["total_multa"] = _fmt(orig["tot"])
        idx_aj = [i for i, m in enumerate(movs) if re.match(r"Ajuste\s+por\s+R\.?\s*Error\s+Material", m["desc"], re.IGNORECASE)]
        if idx_aj:
            acum = movs[: idx_aj[-1] + 1]
            t_mod = sum(m["trib"] for m in acum)
            i_mod = sum(m["int"] for m in acum)
            tot_mod = sum(m["tot"] for m in acum)
            datos["multa_mod"] = _fmt(t_mod)
            datos["interes_multa_mod"] = _fmt(i_mod)
            datos["total_multa_mod"] = _fmt(tot_mod)
            if tot_mod != acum[-1]["sal"]:
                datos["avisos_extraccion"].append("Movimientos de Saldo: el total modificado no coincide con el saldo acumulado; verificar montos.")
        else:
            datos["avisos_extraccion"].append("Movimientos de Saldo: no hay 'Ajuste por R. Error Material'; multa modificada tomada de Eventos.")
    else:
        datos["avisos_extraccion"].append("No se pudo leer Movimientos de Saldo; multa e intereses tomados de otras secciones, verificar.")

    if datos["multa"] == "[COMPLETAR]":
        datos["multa"] = datos["tributo_original"]
        datos["interes_multa"] = datos["interes_original"]
        datos["total_multa"] = datos["total_original"]
    if datos["multa_mod"] == "[COMPLETAR]":
        datos["multa_mod"] = datos["tributo_modificado"]
        datos["interes_multa_mod"] = datos["interes_modificado"]
        datos["total_multa_mod"] = datos["total_modificado"]

    # 11. Pagos con Boleta General
    m_pago = re.search(r"Pago\s+con\s+Boleta\s+General\s+(\d+)\s+(\d+)\s+([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", texto, re.IGNORECASE)
    if m_pago:
        datos["nro_formulario_pago"] = m_pago.group(1).strip()
        datos["nro_boleta_pago"]     = m_pago.group(2).strip()
        datos["pago"]                = "S/ " + m_pago.group(5).strip()
        m_fec_pago = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}\s+Pago\s+con\s+Boleta\s+General", texto, re.IGNORECASE)
        if m_fec_pago:
            datos["fecha_pago"] = m_fec_pago.group(1).strip()

    return datos

# ── Generador de Resolución en Word (Carga y Reemplazo desde Plantilla Seleccionada) ──────────────────
def generar_word(d, ruta_salida, nombre_plantilla="MODELO DJR.docx"):
    plantillas_dir = obtener_directorio_plantillas()
    if not nombre_plantilla:
        nombre_plantilla = "MODELO DJR.docx"
    
    nombre_clean = os.path.basename(nombre_plantilla)
    if not nombre_clean.lower().endswith(".docx"):
        nombre_clean += ".docx"

    ruta_plantilla = os.path.join(plantillas_dir, nombre_clean)
    
    if not os.path.exists(ruta_plantilla):
        if os.path.exists(plantillas_dir):
            for file_item in os.listdir(plantillas_dir):
                if file_item.lower() == nombre_clean.lower():
                    ruta_plantilla = os.path.join(plantillas_dir, file_item)
                    break
                    
    avisos = list(d.get("avisos_extraccion", []))
    if not os.path.exists(ruta_plantilla):
        if os.path.exists(plantillas_dir):
            existentes = [f for f in os.listdir(plantillas_dir) if f.lower().endswith(".docx") and not f.startswith("~$")]
            if existentes:
                ruta_plantilla = os.path.join(plantillas_dir, existentes[0])
                avisos.append(f"No se encontró la plantilla '{nombre_clean}'; se usó '{existentes[0]}'.")

    if not os.path.exists(ruta_plantilla):
        raise FileNotFoundError(f"No se encontró la plantilla Word '{nombre_plantilla}'")

    doc = Document(ruta_plantilla)
    
    # Ajustes según modelo
    # En MODELO INC B y PE, la rectificatoria en los considerandos se refiere al pago de la deuda
    if "inc b" in nombre_clean.lower() or "modelo pe" in nombre_clean.lower():
        fecha_recti = d.get("fecha_pago", "[FECHA PAGO]")
        nro_od_mod = d.get("nro_boleta_pago", "[NRO BOLETA PAGO]")
    else:
        # En MODELO DJR y similares, corresponde a la Declaración Rectificatoria PDT
        f_rec = d.get("fecha_rectificatoria")
        od_rec = d.get("nro_orden_rectificatoria")
        exp_num = d.get("nro_expediente")
        exp_fec = d.get("fecha_expediente")
        fecha_recti = f_rec if (f_rec and f_rec != "[COMPLETAR]" and f_rec != exp_fec) else "[FECHA RECTIFICATORIA]"
        nro_od_mod = od_rec if (od_rec and od_rec != "[COMPLETAR]" and od_rec != exp_num) else "[NRO ORDEN RECTIFICATORIA]"

    fec_dec = d.get("fecha_declaracion")
    if not fec_dec or fec_dec == "[COMPLETAR]":
        fec_dec = "[FECHA DECLARACIÓN]"

    mapeo = {
        "{{NRO_RES}}": d.get("nro_resolucion", "[COMPLETAR]"),
        "{{CIUDAD}}": d.get("ciudad", "Huancayo"),
        "{{FECHA_RES}}": d.get("fecha_resolucion", "[COMPLETAR]"),
        "{{NRO_EXP}}": d.get("nro_expediente", "[COMPLETAR]"),
        "{{FEC_EXP}}": d.get("fecha_expediente", "[COMPLETAR]"),
        "{{RAZON_SOCIAL}}": d.get("razon_social", "[COMPLETAR]"),
        "{{RUC}}": d.get("ruc", "[COMPLETAR]"),
        "{{DOMICILIO}}": d.get("domicilio", "[COMPLETAR]"),
        "{{DOMICILI}}": d.get("domicilio", "[COMPLETAR]"),
        "{{NRO_OP}}": d.get("nro_orden_pago", "[COMPLETAR]"),
        "{{COD_TRIB}}": d.get("codigo_tributo", "[COMPLETAR]"),
        "{{PERIODO}}": d.get("periodo", "[COMPLETAR]"),
        "{{FEC_EM_OP}}": d.get("fecha_emision_op", "[COMPLETAR]"),
        "{{TRIB}}": d.get("tributo_original", "0.00"),
        "{{INT}}": d.get("interes_original", "0.00"),
        "{{TOTAL}}": d.get("total_original", "0.00"),
        "{{COD_TRIB_ASOC}}": d.get("codigo_tributo_asoc", "[COMPLETAR]"),
        "{{MULTA}}": d.get("multa", "0.00"),
        "{{INT_MULTA}}": d.get("interes_multa", "0.00"),
        "{{TOTAL_MULTA}}": d.get("total_multa", "0.00"),
        "{{MULTA_MOD}}": d.get("multa_mod", "0.00"),
        "{{INT_MULTA_MOD}}": d.get("interes_multa_mod", "0.00"),
        "{{TOTAL_MULTA_MOD}}": d.get("total_multa_mod", "0.00"),
        "{{NOM_TRIB_ASOC}}": NOMBRES_TRIBUTO.get(d.get("codigo_tributo_asoc"), "[COMPLETAR]"),
        "{{FEC_INT}}": d.get("fecha_interes", "[COMPLETAR]"),
        "{{FEC_DEC}}": fec_dec,
        "{{NRO_OD}}": d.get("nro_orden_declaracion", "[COMPLETAR]"),
        "{{FECHA_RECTI}}": fecha_recti,
        "{{NRO_OD_MOD}}": nro_od_mod,
        "{{TRIB_MOD}}": d.get("tributo_modificado", "0.00"),
        "{{INT_MOD}}": d.get("interes_modificado", "0.00"),
        "{{TOTAL_MOD}}": d.get("total_modificado", "0.00"),
        "{{PAGO}}": d.get("pago", "S/ 0.00"),
        "{{FECHA_PE }}": d.get("fecha_expediente", "[COMPLETAR]"),
        "{{RI_PE}}": d.get("nro_resolucion", "[COMPLETAR]"),
        "{{URD}}": d.get("nro_expediente", "[COMPLETAR]"),
        "{{INTENDENCIA}}": d.get("intendencia", "[COMPLETAR]")
    }
    
    def consolidar_runs_parrafo(p):
        runs_text = [r.text for r in p.runs]
        text_completo = "".join(runs_text)
        
        if "{{" not in text_completo or "}}" not in text_completo:
            return
            
        for match in re.finditer(r"\{\{[^}]+\}\}", text_completo):
            start_idx = match.start()
            end_idx = match.end()
            
            char_count = 0
            first_run_idx = -1
            last_run_idx = -1
            
            for i, r in enumerate(p.runs):
                run_len = len(r.text) if r.text else 0
                run_start = char_count
                run_end = char_count + run_len
                
                if run_start <= start_idx < run_end:
                    first_run_idx = i
                if run_start < end_idx <= run_end:
                    last_run_idx = i
                    break
                    
                char_count += run_len
                
            if first_run_idx != -1 and last_run_idx != -1:
                merged_text = "".join(p.runs[j].text for j in range(first_run_idx, last_run_idx + 1) if p.runs[j].text)
                p.runs[first_run_idx].text = merged_text
                for j in range(first_run_idx + 1, last_run_idx + 1):
                    p.runs[j].text = ""
                    
    def procesar_parrafo(p):
        consolidar_runs_parrafo(p)
        texto = p.text
        if not texto or "{{" not in texto:
            return
            
        for key, val in mapeo.items():
            if key in texto:
                for run in p.runs:
                    if run.text and key in run.text:
                        run.text = run.text.replace(key, str(val))
                        
        # Sustitución de seguridad para llaves que queden fragmentadas
        if "{{" in p.text:
            full_text = "".join(r.text for r in p.runs)
            modificado = False
            for key, val in mapeo.items():
                if key in full_text:
                    full_text = full_text.replace(key, str(val))
                    modificado = True
            if modificado:
                for r in p.runs:
                    r.text = ""
                if p.runs:
                    p.runs[0].text = full_text
                else:
                    p.add_run(full_text)
                        
    # Procesar párrafos
    for p in doc.paragraphs:
        procesar_parrafo(p)
        
    # Procesar tablas (cuerpo del documento y tablas de considerandos/resolutivas)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    procesar_parrafo(p)
                    
    # Procesar Headers y Footers en todas las secciones
    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                procesar_parrafo(p)
            for table in section.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            procesar_parrafo(p)
        if section.footer:
            for p in section.footer.paragraphs:
                procesar_parrafo(p)
            for table in section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            procesar_parrafo(p)
                            
    # Avisos: campos usados por la plantilla que quedaron sin dato o llaves desconocidas
    textos = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                textos.extend(p.text for p in cell.paragraphs)
    texto_final = "\n".join(textos)
    pendientes = sorted(set(re.findall(r"\[[A-ZÁÉÍÓÚ ]{4,}\]", texto_final)))
    if pendientes:
        avisos.append("Datos no encontrados en el PDF: " + ", ".join(pendientes) + ".")
    sin_reemplazar = sorted(set(re.findall(r"\{\{[^}]*\}\}", texto_final)))
    if sin_reemplazar:
        avisos.append("Campos de la plantilla sin reemplazar: " + ", ".join(sin_reemplazar) + ".")

    doc.save(ruta_salida)
    return avisos

# ── Endpoints de Gestión de Plantillas / Modelos Word ────────────────────────────

@router.get("/api/models")
async def list_models():
    """Devuelve la lista de plantillas .docx disponibles."""
    plantillas_dir = obtener_directorio_plantillas()
    modelos = []
    if os.path.exists(plantillas_dir):
        for file in os.listdir(plantillas_dir):
            if file.lower().endswith(".docx") and not file.startswith("~$"):
                full_path = os.path.join(plantillas_dir, file)
                size_bytes = os.path.getsize(full_path)
                mtime = os.path.getmtime(full_path)
                nombre_display = os.path.splitext(file)[0]
                modelos.append({
                    "filename": file,
                    "name": nombre_display,
                    "size": size_bytes,
                    "mtime": mtime,
                    "is_default": file.lower() in ["modelo.docx", "modelo djr.docx"]
                })
    modelos.sort(key=lambda x: x["name"].lower())
    return {"status": "success", "models": modelos}

@router.post("/api/open-templates-folder")
async def open_templates_folder():
    """Abre la carpeta física de plantillas Word en el Explorador de Windows."""
    try:
        plantillas_dir = obtener_directorio_plantillas()
        if os.name == 'nt' and os.path.exists(plantillas_dir):
            os.startfile(plantillas_dir)
        return {"status": "success", "path": plantillas_dir}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})

@router.post("/api/upload-model")
async def upload_model(files: List[UploadFile] = File(...)):
    """Sube una o varias plantillas .docx al directorio de plantillas de RIbot."""
    if not files:
        return JSONResponse(status_code=400, content={"status": "error", "message": "No se enviaron archivos."})
    
    plantillas_dir = obtener_directorio_plantillas()
    os.makedirs(plantillas_dir, exist_ok=True)
    subidos = []
    for file in files:
        if not file.filename.lower().endswith(".docx"):
            continue
        nombre_limpio = sanitizar_nombre_archivo(os.path.splitext(file.filename)[0]) + ".docx"
        dest_path = os.path.join(plantillas_dir, nombre_limpio)
        with open(dest_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        subidos.append(nombre_limpio)
        
    if not subidos:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Formato inválido. Solo se admiten plantillas .docx"})
        
    return JSONResponse(content={"status": "success", "message": f"Se agregaron {len(subidos)} plantilla(s) correctamente.", "uploaded": subidos})

@router.delete("/api/delete-model/{filename}")
async def delete_model(filename: str):
    """Elimina una plantilla personalizada .docx."""
    safe_name = os.path.basename(filename)
    if safe_name == "MODELO.docx":
        return JSONResponse(status_code=400, content={"status": "error", "message": "No se puede eliminar la plantilla base MODELO.docx."})
    plantillas_dir = obtener_directorio_plantillas()
    dest_path = os.path.join(plantillas_dir, safe_name)
    if os.path.exists(dest_path):
        os.remove(dest_path)
        return JSONResponse(content={"status": "success", "message": f"Plantilla '{safe_name}' eliminada correctamente."})
    return JSONResponse(status_code=404, content={"status": "error", "message": "Plantilla no encontrada."})

# ── Endpoints de la API del Módulo RIbot ───────────────────────────────────────────

@router.get("/")
async def ribot_home(request: Request):
    limpiar_archivos_temporales()
    return templates.TemplateResponse("ribot/index.html", {
        "request": request,
        "modulo": "RIbot"
    })

@router.post("/api/process-upload")
async def process_upload(
    files: List[UploadFile] = File(...),
    destination_path: str = Form(""),
    model_name: str = Form("MODELO.docx"),
    file_models: str = Form("{}")
):
    """Sube PDFs de Fichas de Valor, genera Word y guarda en destino físico y/o descarga."""
    limpiar_archivos_temporales()

    if not files:
        return JSONResponse(status_code=400, content={"status": "error", "message": "No se subieron archivos."})
        
    dest_dir = None
    if destination_path.strip():
        try:
            dest_dir = os.path.abspath(destination_path.strip())
            os.makedirs(dest_dir, exist_ok=True)
        except Exception as e:
            return JSONResponse(status_code=400, content={"status": "error", "message": f"Ruta de destino inválida o sin permisos: {e}"})

    job_id = str(uuid.uuid4())
    job_path = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_path, exist_ok=True)
    
    input_path = os.path.join(job_path, "input")
    os.makedirs(input_path, exist_ok=True)
    
    generated_files = []
    avisos = []
    import json
    mapping_modelos = {}
    try:
        if file_models and file_models.strip():
            mapping_modelos = json.loads(file_models)
    except Exception:
        mapping_modelos = {}

    for file in files:
        if not file.filename.lower().endswith(".pdf"):
            avisos.append(f"{file.filename}: se omitió (no es PDF).")
            continue
            
        safe_pdf_name = sanitizar_nombre_archivo(os.path.basename(file.filename)) + ".pdf"
        pdf_path = os.path.join(input_path, safe_pdf_name)
        
        with open(pdf_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
            
        # Extracción y Redacción
        d = extraer_datos_pdf(pdf_path)
        
        # Obtener el modelo asignado para este PDF (o el global)
        modelo_especifico = mapping_modelos.get(file.filename, model_name)
        if not modelo_especifico:
            modelo_especifico = model_name or "MODELO.docx"

        ruc_safe = sanitizar_nombre_archivo(d.get("ruc", "SIN_RUC"))
        nro_res_safe = sanitizar_nombre_archivo(d.get("nro_resolucion", "SIN_NRO"))
        razon_safe = sanitizar_nombre_archivo(d.get("razon_social", "SIN_NOMBRE"))
        
        nombre_word = f"{ruc_safe}_{nro_res_safe}_{razon_safe}_mod.docx"
        ruta_word = os.path.join(job_path, nombre_word)
        
        try:
            for av in generar_word(d, ruta_word, nombre_plantilla=modelo_especifico):
                avisos.append(f"{file.filename}: {av}")
            generated_files.append({
                "filename": nombre_word,
                "path": ruta_word
            })
            
            if dest_dir:
                ruta_destino_fisico = os.path.join(dest_dir, nombre_word)
                shutil.copy2(ruta_word, ruta_destino_fisico)
        except Exception as e:
            print(f"Error redactando Word {nombre_word}: {e}")
            avisos.append(f"{file.filename}: no se pudo generar el Word ({e}).")
            
    if not generated_files:
        return JSONResponse(status_code=400, content={"status": "error", "message": "No se pudieron procesar las Fichas del Valor PDF subidas."})
        
    if len(generated_files) == 1:
        filename = generated_files[0]["filename"]
        download_url = f"/ribot/api/download/{job_id}/{filename}"
        return JSONResponse(content={
            "status": "success",
            "type": "single",
            "filename": filename,
            "download_url": download_url,
            "count": 1,
            "physical_path": dest_dir,
            "saved_locally": bool(dest_dir),
            "avisos": avisos
        })
    else:
        zip_filename = f"Resoluciones_SUNAT_{job_id[:8]}.zip"
        zip_path = os.path.join(job_path, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for item in generated_files:
                zipf.write(item["path"], item["filename"])
                
        download_url = f"/ribot/api/download/{job_id}/{zip_filename}"
        return JSONResponse(content={
            "status": "success",
            "type": "massive",
            "filename": zip_filename,
            "download_url": download_url,
            "count": len(generated_files),
            "physical_path": dest_dir,
            "saved_locally": bool(dest_dir),
            "avisos": avisos
        })

@router.post("/api/process-folder")
async def process_folder(data: dict):
    """Procesa una carpeta del servidor que contiene Fichas de Valor PDF y guarda en destino físico."""
    limpiar_archivos_temporales()

    carpeta = data.get("carpeta", "").strip()
    destination_path = data.get("destination_path", "").strip()
    model_name = data.get("model_name", "MODELO.docx")
    mapping_modelos = data.get("file_models", {})
    
    is_absolute = os.path.isabs(carpeta) or (len(carpeta) >= 2 and carpeta[1] == ':')
    
    try:
        if is_absolute:
            target_path = os.path.abspath(carpeta)
        else:
            target_path = os.path.abspath(os.path.join(RIBOT_BASE_DIR, carpeta.strip(r'\/')))
            if not target_path.startswith(os.path.abspath(RIBOT_BASE_DIR)):
                 return JSONResponse(status_code=400, content={"status": "error", "message": "Acceso denegado a esta ruta"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"status": "error", "message": f"Ruta inválida: {e}"})
             
    if not os.path.isdir(target_path):
        return JSONResponse(status_code=400, content={"status": "error", "message": "La carpeta de origen no existe en el servidor."})
        
    try:
        pdfs = [f for f in os.listdir(target_path) if f.lower().endswith(".pdf")]
    except Exception as e:
        return JSONResponse(status_code=400, content={"status": "error", "message": f"Sin permisos o error al leer la carpeta: {e}"})

    if not pdfs:
        return JSONResponse(status_code=400, content={"status": "error", "message": "No se encontraron PDFs de Fichas de Valor en la carpeta seleccionada."})
        
    dest_dir = None
    if destination_path.strip():
        try:
            dest_dir = os.path.abspath(destination_path.strip())
            os.makedirs(dest_dir, exist_ok=True)
        except Exception as e:
            return JSONResponse(status_code=400, content={"status": "error", "message": f"Ruta de destino inválida o sin permisos: {e}"})

    job_id = str(uuid.uuid4())
    job_path = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_path, exist_ok=True)
    
    generated_files = []
    avisos = []
    
    for pdf_name in pdfs:
        pdf_path = os.path.join(target_path, pdf_name)
        d = extraer_datos_pdf(pdf_path)
        
        modelo_pdf = mapping_modelos.get(pdf_name, model_name)
        if not modelo_pdf:
            modelo_pdf = model_name or "MODELO.docx"

        ruc_safe = sanitizar_nombre_archivo(d.get("ruc", "SIN_RUC"))
        nro_res_safe = sanitizar_nombre_archivo(d.get("nro_resolucion", "SIN_NRO"))
        razon_safe = sanitizar_nombre_archivo(d.get("razon_social", "SIN_NOMBRE"))
        
        nombre_word = f"{ruc_safe}_{nro_res_safe}_{razon_safe}_mod.docx"
        
        if dest_dir:
            ruta_word = os.path.join(dest_dir, nombre_word)
        else:
            ruta_word = os.path.join(target_path, nombre_word)
            
        try:
            for av in generar_word(d, ruta_word, nombre_plantilla=modelo_pdf):
                avisos.append(f"{pdf_name}: {av}")
            generated_files.append({
                "filename": nombre_word,
                "path": ruta_word
            })
        except Exception as e:
            print(f"Error en procesamiento local de {nombre_word}: {e}")
            avisos.append(f"{pdf_name}: no se pudo generar el Word ({e}).")
            
    if not generated_files:
        return JSONResponse(status_code=500, content={"status": "error", "message": "No se pudieron redactar las resoluciones de la carpeta."})
        
    zip_filename = f"Resoluciones_Local_{job_id[:8]}.zip"
    zip_path = os.path.join(job_path, zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for item in generated_files:
            zipf.write(item["path"], item["filename"])
            
    download_url = f"/ribot/api/download/{job_id}/{zip_filename}"
    return JSONResponse(content={
        "status": "success",
        "type": "massive",
        "filename": zip_filename,
        "download_url": download_url,
        "count": len(generated_files),
        "physical_path": dest_dir if dest_dir else target_path,
        "saved_locally": True,
        "avisos": avisos
    })

@router.get("/api/download/{job_id}/{filename}")
async def download_results(job_id: str, filename: str):
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(TEMP_DIR, job_id, safe_filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=safe_filename)
    return JSONResponse(status_code=404, content={"message": "Archivo no encontrado"})

@router.get("/api/list-folders")
async def list_folders(path: str = ""):
    """Lista subcarpetas locales. Soporta rutas absolutas locales o relativas a ribot_docs."""
    try:
        path_str = path.strip()
        is_absolute = False
        
        if os.path.isabs(path_str) or (len(path_str) >= 2 and path_str[1] == ':'):
            is_absolute = True
            target_path = os.path.abspath(path_str)
        else:
            target_path = os.path.abspath(os.path.join(RIBOT_BASE_DIR, path_str.strip(r'\/')))
            if not target_path.startswith(os.path.abspath(RIBOT_BASE_DIR)):
                return {"folders": [], "error": "Acceso denegado"}
            
        if not os.path.isdir(target_path):
            return {"folders": [], "error": "La ruta especificada no existe o no es un directorio válido."}
        
        folders = []
        try:
            for item in os.listdir(target_path):
                full_path = os.path.join(target_path, item)
                if os.path.isdir(full_path):
                    if is_absolute:
                        folders.append({
                            "name": item,
                            "path": full_path
                        })
                    else:
                        rel_path = os.path.relpath(full_path, RIBOT_BASE_DIR)
                        folders.append({
                            "name": item,
                            "path": rel_path
                        })
        except PermissionError:
            return {"folders": [], "error": "Sin permisos de acceso en esta carpeta."}
        
        folders.sort(key=lambda x: x["name"].lower())
        
        current_display = target_path if is_absolute else (os.path.relpath(target_path, RIBOT_BASE_DIR) if target_path != os.path.abspath(RIBOT_BASE_DIR) else "")
        
        return {"folders": folders, "current": current_display, "is_absolute": is_absolute}
        
    except Exception as e:
        return {"folders": [], "error": str(e)}
