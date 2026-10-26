import os
import cv2
import fitz  # PyMuPDF
import pandas as pd
from pdf2image import convert_from_path
import shutil
import unicodedata
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import List, Dict, Any, Optional, Callable

# Textos repetidos
_AUXILIAR = "auxiliar"
_FILENAME = "filename"
_FIRMADO = "Firmado"
_FIRMANTE = "Firmante"
_FIRMANTE_2 = "firmante"
_STATUS = "status"
_VALID = "valid"

# TODO: Make these paths configurable via settings or environment variables
POPPLER_PATH = r"C:\poppler\Library\bin"
DEFAULT_SIGNATURES_DIR = r"D:\FIRMAS_AUXILIARES_COACTIVOS"


def quitar_tildes(texto: str) -> str:
    """Normaliza texto: quita tildes y acentos."""
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def extraer_zonas_de_firma(pdf_path: str, salida_path: str, divisiones: int = 3):
    """Extrae la última página del PDF y la divide en zonas para buscar firmas."""
    # Convert only the last page to image
    # We use poppler_path if it exists, otherwise hope it's in system PATH
    kwargs = {}
    if os.path.exists(POPPLER_PATH):
        kwargs["poppler_path"] = POPPLER_PATH

    # Let's use PyMuPDF to get total pages first
    try:
        with fitz.open(pdf_path) as doc:
            total_pages = doc.page_count

        paginas = convert_from_path(pdf_path, first_page=total_pages, last_page=total_pages, **kwargs)
    except Exception as e:
        print(f"Error convirtiendo PDF {pdf_path} a imagen: {e}")
        return [], []

    if not paginas:
        return [], []

    ultima_pagina = paginas[0]
    ancho, alto = ultima_pagina.size
    nombre_base = os.path.splitext(os.path.basename(pdf_path))[0]
    zona_alto = alto // divisiones
    zonas_extraidas = []

    for i in range(divisiones):
        y_inicio = i * zona_alto
        y_fin = (i + 1) * zona_alto if i < divisiones - 1 else alto
        zona = ultima_pagina.crop((0, y_inicio, ancho, y_fin))
        ruta_salida = os.path.join(salida_path, f"{nombre_base}_zona_{i + 1}.png")
        zona.save(ruta_salida)
        zonas_extraidas.append(ruta_salida)

    return zonas_extraidas, paginas


def _mejor_coincidencia(firma_ref, zonas: List[str], umbral: float) -> Optional[float]:
    """Similitud de la primera zona que alcanza el umbral con la firma de referencia, o None."""
    for zona_path in zonas:
        zona = cv2.imread(zona_path, cv2.IMREAD_GRAYSCALE)
        if zona is None:
            continue
        # La zona no puede ser más pequeña que la referencia
        if zona.shape[0] < firma_ref.shape[0] or zona.shape[1] < firma_ref.shape[1]:
            continue
        resultado = cv2.matchTemplate(zona, firma_ref, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(resultado)
        if max_val >= umbral:
            return max_val
    return None


def comparar_firmas(zonas: List[str], carpeta_firmas: str, umbral: float = 0.7) -> Dict[str, str]:
    """Compara las zonas extraídas con las firmas de referencia en la carpeta."""
    if not os.path.isdir(carpeta_firmas):
        return {_FIRMADO: "Error", _FIRMANTE: "Carpeta de firmas no encontrada"}

    firmas = [f for f in os.listdir(carpeta_firmas) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
    coincidencias = []

    for firma_archivo in firmas:
        nombre_firma = os.path.splitext(firma_archivo)[0].split("_")[0].strip()
        firma_ref = cv2.imread(os.path.join(carpeta_firmas, firma_archivo), cv2.IMREAD_GRAYSCALE)
        if firma_ref is None:
            continue
        max_val = _mejor_coincidencia(firma_ref, zonas, umbral)
        if max_val is not None:
            coincidencias.append((nombre_firma, max_val))

    if coincidencias:
        mejor = max(coincidencias, key=lambda x: x[1])
        return {_FIRMADO: "Sí", _FIRMANTE: mejor[0]}
    return {_FIRMADO: "No", _FIRMANTE: ""}


def buscar_auxiliar_en_pdf(pdf_path: str, ruta_txt_auxiliares: str) -> str:
    """Busca nombres de auxiliares en el texto del PDF."""
    if not os.path.exists(ruta_txt_auxiliares):
        return "LISTA NO ENCONTRADA"

    with open(ruta_txt_auxiliares, "r", encoding="utf-8") as f:
        auxiliares = [quitar_tildes(line.strip().upper()) for line in f if line.strip()]

    texto_completo = ""
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                texto_completo += page.get_text().upper()
        texto_completo = quitar_tildes(texto_completo)
    except Exception as e:
        print(f"Error extrayendo texto de {pdf_path}: {e}")
        return "ERROR LECTURA"

    mejor_match = ""
    max_palabras = 0

    for nombre in auxiliares:
        palabras = nombre.split()
        # Cuenta cuántas palabras del nombre están en el texto del PDF
        coincidencias = sum(1 for palabra in palabras if palabra in texto_completo)
        if coincidencias > max_palabras and coincidencias >= 2:  # al menos 2 palabras
            mejor_match = nombre
            max_palabras = coincidencias

    return mejor_match if mejor_match else "SIN CONSIGNA"


def _resultado_pdf(nombre_pdf: str, firmante: str, firmado: bool, status: str, auxiliar: str) -> Dict[str, Any]:
    return {
        _FILENAME: nombre_pdf,
        _FIRMANTE_2: firmante,
        _STATUS: status,
        _VALID: firmado,
        _AUXILIAR: auxiliar,
        "expediente": "---",
        "ruc": "---",
    }


def _procesar_pdf(archivo, carpeta_pdfs, carpeta_temp, carpeta_firmas, ruta_txt_auxiliares) -> Dict[str, Any]:
    ruta_pdf = os.path.join(carpeta_pdfs, archivo)
    try:
        # 1. Extraer zonas de firma, 2. compararlas con las firmas de referencia, 3. buscar auxiliar en texto
        zonas, _ = extraer_zonas_de_firma(ruta_pdf, carpeta_temp)
        res_firmas = comparar_firmas(zonas, carpeta_firmas)
        auxiliar = buscar_auxiliar_en_pdf(ruta_pdf, ruta_txt_auxiliares)

        firmado = res_firmas[_FIRMADO] == "Sí"
        return _resultado_pdf(
            archivo, res_firmas[_FIRMANTE], firmado, "Firma Válida" if firmado else "Sin Firma", auxiliar
        )
    except Exception as e:
        print(f"⚠️ Error en {archivo}: {e}")
        return _resultado_pdf(archivo, "ERROR", False, "Error", "ERROR")


def _autoajustar_columnas(ws) -> None:
    for col in ws.columns:
        max_length = 0
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = max_length + 2


def _guardar_reporte_excel(resultados: List[Dict[str, Any]], ruta_excel: str) -> None:
    df = pd.DataFrame(
        [
            {
                "Archivo": r[_FILENAME],
                _FIRMADO: "SÍ" if r[_VALID] else "NO",
                _FIRMANTE: r[_FIRMANTE_2],
                "Auxiliar Coactivo": r[_AUXILIAR],
            }
            for r in resultados
        ]
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REPORTE"
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    if not df.empty:
        ultima_col = openpyxl.utils.get_column_letter(df.shape[1])
        tabla = Table(displayName="TablaSigna", ref=f"A1:{ultima_col}{df.shape[0] + 1}")
        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tabla)
        _autoajustar_columnas(ws)

    wb.save(ruta_excel)


def procesar_carpeta(
    carpeta_pdfs: str,
    carpeta_firmas: str = DEFAULT_SIGNATURES_DIR,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
) -> Dict[str, Any]:
    """Procesa todos los PDFs en una carpeta y genera un reporte."""

    carpeta_temp = os.path.join(carpeta_pdfs, "__temp_zonas__")
    os.makedirs(carpeta_temp, exist_ok=True)

    archivos_pdf = [f for f in os.listdir(carpeta_pdfs) if f.lower().endswith(".pdf")]
    total = len(archivos_pdf)
    ruta_txt_auxiliares = os.path.join(carpeta_firmas, "AUXILIARESCOACTIVOS.txt")

    resultados = []
    for i, archivo in enumerate(archivos_pdf):
        if on_progress:
            on_progress(i + 1, total, archivo)
        resultados.append(_procesar_pdf(archivo, carpeta_pdfs, carpeta_temp, carpeta_firmas, ruta_txt_auxiliares))

    if os.path.exists(carpeta_temp):
        shutil.rmtree(carpeta_temp)

    reporte_filename = "REPORTE_SIGNA_FIRMAS.xlsx"
    _guardar_reporte_excel(resultados, os.path.join(carpeta_pdfs, reporte_filename))

    return {
        _STATUS: "success",
        "data": resultados,
        "reporte_excel": reporte_filename,  # Return just the filename, router will handle serving it
        "summary": {
            "total": total,
            _VALID: len([r for r in resultados if r[_VALID]]),
            "error": len([r for r in resultados if not r[_VALID]]),
        },
    }
