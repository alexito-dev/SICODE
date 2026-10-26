import os
import re
import io
import base64
import logging
from typing import List, Dict, Any, Optional
import cv2
import numpy as np
from PIL import Image, ImageOps
import pytesseract
import qrcode

# --- Configuración de Logs ---
logger = logging.getLogger(__name__)

# --- Configuración Tesseract ---
tesseract_paths = [
    "/usr/bin/tesseract",  # Linux / PythonAnywhere
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\alets\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
]

tesseract_cmd = None
for path in tesseract_paths:
    if os.path.exists(path):
        tesseract_cmd = path
        break

if tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    logger.info(f"Tesseract encontrado en: {tesseract_cmd}")
else:
    logger.warning("Tesseract NO encontrado. El OCR fallará.")


def clean_ocr_number(text: str) -> str:
    replacements = {
        "O": "0",
        "o": "0",
        "Q": "0",
        "D": "0",
        "C": "0",
        "I": "1",
        "l": "1",
        "|": "1",
        "!": "1",
        "i": "1",
        "L": "1",
        "j": "1",
        "J": "1",
        "]": "1",
        "[": "1",
        "Z": "2",
        "E": "3",
        "A": "4",
        "S": "5",
        "$": "5",
        "G": "6",
        "T": "7",
        "B": "8",
        "g": "9",
    }
    cleaned = text
    for char, digit in replacements.items():
        cleaned = cleaned.replace(char, digit)
    return re.sub(r"\D", "", cleaned)


def validate_ruc(ruc: str) -> bool:
    if len(ruc) != 11 or not ruc.isdigit():
        return False
    factors = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(ruc[i]) * factors[i] for i in range(10))
    residuo = suma % 11
    complemento = 11 - residuo
    digito = 0 if complemento == 10 else (1 if complemento == 11 else complemento)
    return digito == int(ruc[10])


def preprocess_image(pil_image):
    open_cv_image = np.array(pil_image)
    img_gray = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2GRAY) if len(open_cv_image.shape) == 3 else open_cv_image

    # Median Blur suave
    img_blur = cv2.medianBlur(img_gray, 3)

    # Adaptive Threshold
    img_thresh = cv2.adaptiveThreshold(img_blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 15)
    return Image.fromarray(img_thresh), img_thresh


MESES_MAP = {
    "ENE": "01",
    "FEB": "02",
    "MAR": "03",
    "ABR": "04",
    "MAY": "05",
    "JUN": "06",
    "JUL": "07",
    "AGO": "08",
    "SET": "09",
    "SEP": "09",
    "OCT": "10",
    "NOV": "11",
    "DIC": "12",
    "enero": "01",
    "febrero": "02",
    "marzo": "03",
    "abril": "04",
    "mayo": "05",
    "junio": "06",
    "julio": "07",
    "agosto": "08",
    "septiembre": "09",
    "setiembre": "09",
    "octubre": "10",
    "noviembre": "11",
    "diciembre": "12",
}

_OCR_DIGITS = r"[\dOIZSBl|!i]"


def _extract_exp_sigad(text: str, data: Dict[str, str]) -> None:
    match = re.search(r"(\d{3}-[\w\d]+-\d{4}-\d+-\d)", text)
    if match:
        data["exp_sigad"] = match.group(1)


def _extract_res_coactiva(text: str, data: Dict[str, str]) -> None:
    pattern = r"(?i)(?:RESOLUCI[ÓO]N|RES\.?)\s*(?:COACTIVA)?\s*(?:N[º°])?[\s\|i:\.\'‘’]*([\dOIZSBl|!i]{10,16})"
    match = re.search(pattern, text)
    if match:
        raw = clean_ocr_number(match.group(1))
        if len(raw) >= 10:
            data["res_coactiva"] = raw
    if "res_coactiva" not in data:
        possibles = re.findall(r"\b(133\d{10})\b", clean_ocr_number(text))
        if possibles:
            data["res_coactiva"] = possibles[0]


def _extract_expediente_rc(text: str, data: Dict[str, str]) -> None:
    pattern = r"(?i)EXPEDIENTE\s*(?:NUMERO|N[º°])?[\s\|i:\.\'‘’]*([\dOIZSBl|!i]{10,16})"
    match = re.search(pattern, text)
    if match:
        data["expediente_rc"] = clean_ocr_number(match.group(1))


def _assign_ruc(data: Dict[str, str], ruc: str, context_area: str) -> None:
    if any(w in context_area for w in ["USUARIO", "TERCERO", "EJECUTORA"]):
        data["ruc_tercero"] = ruc
    elif not data.get("ruc_contribuyente"):
        data["ruc_contribuyente"] = ruc
    elif ruc != data.get("ruc_contribuyente") and not data.get("ruc_tercero"):
        data["ruc_tercero"] = ruc


def _extract_rucs(text: str, data: Dict[str, str]) -> None:
    for match in re.finditer(r"\b([\dOIZSBl|!i]{11})\b", text):
        ruc = clean_ocr_number(match.group(1))
        if validate_ruc(ruc):
            start_pos = max(0, match.start() - 100)
            _assign_ruc(data, ruc, text.upper()[start_pos : match.start()])


def _clean_nombre_contribuyente(val: str) -> str:
    val_clean = re.sub(r"(?i)^(?:DEUDOR|TRIBUTARIO|CONTRIBUYENTE|CONTRIB)[:\s\|!\"\'\.\-=]*", "", val).strip()
    return re.sub(r"^[^\w\(\)]+", "", val_clean).strip()


def _extract_nombre_contribuyente(lines: List[str], data: Dict[str, str]) -> None:
    for i, line in enumerate(lines):
        if not any(w in line.upper() for w in ["DEUDOR", "CONTRIBUYENTE", "CONTRIB:"]):
            continue
        match = re.search(r"[:\.]\s*(.*)", line)
        val = ""
        if match:
            val = match.group(1).strip()
        elif i + 1 < len(lines):
            val = lines[i + 1].strip()
        if val:
            val_clean = _clean_nombre_contribuyente(val)
            if len(val_clean) > 4:
                data["nombre_contribuyente"] = val_clean
        break


def _extract_nombre_tercero(text: str, data: Dict[str, str]) -> None:
    match = re.search(r"(?i)USUARIO:\s*(?:RUC)?\s*(\d{11})?\s*(.*)", text)
    if not match:
        return
    if match.group(1):
        data["ruc_tercero"] = match.group(1)
    name_clean = re.split(r"(?i)PROC:|ASUNTO:", match.group(2).strip())[0].strip()
    name_clean = re.sub(r"^[^\w]+", "", name_clean).strip()
    if len(name_clean) > 4:
        data["nombre_tercero"] = name_clean


def _format_monto(raw_amount: str) -> Optional[str]:
    raw = clean_ocr_number(raw_amount)
    if len(raw) < 3:
        return None
    val = float(int(raw[:-2]) + int(raw[-2:]) / 100.0)
    return f"S/. {val:,.2f}"


def _extract_monto(text: str, data: Dict[str, str]) -> None:
    monto_paren = re.search(r"\(\s*([\dOIZSB]{1,10}[.,]\d{2})\s*\)", text)
    if monto_paren:
        monto = _format_monto(monto_paren.group(1))
        if monto:
            data["monto"] = monto
    monto_soles = re.search(r"(?:S/|Soles|suma de)\.?\s*.*?([\d\.,]{1,15}[.,]\d{2})", text, re.I)
    if monto_soles and "monto" not in data:
        monto = _format_monto(monto_soles.group(1))
        if monto:
            data["monto"] = monto


def _extract_fecha_sello(text: str, data: Dict[str, str]) -> None:
    pattern = r"([\dOIZSBl|!i\]\[]{1,4})\s+([A-Z]{3,10})[.,\s‘]*\s*(20\d{2})"
    for m_day, m_mes, m_anio in re.findall(pattern, text.upper()):
        dia = clean_ocr_number(m_day).zfill(2)
        mes = MESES_MAP.get(m_mes) or MESES_MAP.get(m_mes[:3])
        if mes:
            data["fecha_recepcion"] = f"STAMP:{dia}/{mes}/{m_anio}"
            break


def _extract_fechas_numericas(text: str, data: Dict[str, str]) -> None:
    for m in re.finditer(r"\b(\d{1,2})[/-](\d{1,2})[/-](20\d{2})\b", text):
        pre = text[max(0, m.start() - 30) : m.start()].upper()
        if any(w in pre for w in ["FECHA", "HORA", "PRINT"]):
            continue
        f_num = f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"
        if "fecha_recepcion" not in data:
            data["fecha_recepcion"] = f"NUMERIC:{f_num}"
        elif "fecha_rc" not in data and f_num not in data["fecha_recepcion"]:
            data["fecha_rc"] = f_num


def _extract_fecha_texto(text: str, data: Dict[str, str]) -> None:
    meses_pattern = (
        r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)"
    )
    match = re.search(
        rf"(?:\w+[,\s]+)?(\d{{1,2}})\s*[‘\w]{{1,3}}\s+{meses_pattern}\s*[‘\w]{{1,4}}\s+(20\d{{2}})", text, re.IGNORECASE
    )
    if not match:
        return
    d, m_txt, y = match.groups()
    if m_txt.lower() in MESES_MAP:
        val_text = f"{d.zfill(2)}/{MESES_MAP[m_txt.lower()]}/{y}"
        data["fecha_rc"] = val_text
        if "fecha_recepcion" not in data or "NUMERIC:" in data["fecha_recepcion"]:
            data["fecha_recepcion"] = val_text


def _extract_banco(text: str, data: Dict[str, str]) -> None:
    bancos = ["NACION", "CREDITO", "BCP", "INTERBANK", "BBVA", "CONTINENTAL", "SCOTIABANK", "BANBIF"]
    for banco in bancos:
        if banco in text.upper():
            data["banco_cheque"] = banco
            break


def extract_data_from_text(text: str) -> Dict[str, str]:
    data: Dict[str, str] = {}
    text_clean = re.sub(r"  +", " ", text)
    lines = [ln.strip() for ln in text_clean.split("\n") if ln.strip()]

    _extract_exp_sigad(text, data)
    _extract_res_coactiva(text, data)
    _extract_expediente_rc(text, data)
    _extract_rucs(text, data)
    _extract_nombre_contribuyente(lines, data)
    _extract_nombre_tercero(text, data)
    _extract_monto(text, data)
    _extract_fecha_sello(text, data)
    _extract_fechas_numericas(text, data)
    _extract_fecha_texto(text, data)
    _extract_banco(text, data)
    return data


def _empty_fields() -> Dict[str, str]:
    return {
        "exp_sigad": "",
        "fecha_recepcion": "",
        "ruc_contribuyente": "",
        "nombre_contribuyente": "",
        "res_coactiva": "",
        "fecha_rc": "",
        "expediente_rc": "",
        "monto": "",
        "ruc_tercero": "",
        "nombre_tercero": "",
        "cheque_boleta": "",
        "banco_cheque": "",
        "observaciones": "",
    }


def _score_ocr_data(data: Dict[str, str]) -> int:
    pesos = {"exp_sigad": 5, "ruc_contribuyente": 4, "res_coactiva": 4, "fecha_recepcion": 5}
    return sum(peso for campo, peso in pesos.items() if campo in data) + len(data)


def _best_rotation_ocr(processed_pil, idx: int):
    """Prueba cuatro rotaciones y devuelve (datos, texto crudo) de la mejor."""
    best_img_data = {}
    max_img_score = -1
    current_raw = f"[Página {idx + 1} | No se detectó texto]"

    for angle in [0, 180, 270, 90]:
        img_to_process = processed_pil if angle == 0 else processed_pil.rotate(angle, expand=True)
        text = pytesseract.image_to_string(img_to_process, config="--oem 3 --psm 6")
        data = extract_data_from_text(text)
        score = _score_ocr_data(data)

        if score > max_img_score:
            max_img_score = score
            best_img_data = data
            current_raw = f"[Página {idx + 1} | Rotación {angle}°]\n" + text

        if score >= 15:
            break
    return best_img_data, current_raw


def _merge_field(merged: Dict[str, str], key: str, value: str) -> None:
    current = merged[key]
    if key == "fecha_recepcion":
        if not current or "STAMP:" in value and "STAMP:" not in current:
            merged[key] = value
    elif not current:
        merged[key] = value
    elif key in ["monto", "exp_sigad", "res_coactiva"]:
        if len(value) >= len(current):
            merged[key] = value
    elif len(value) > len(current):
        merged[key] = value


async def process_ocr_batch(images_b64: List[str]) -> Dict[str, Any]:
    merged_fields = _empty_fields()
    processed_images = []
    all_raw_text = []
    all_results = []

    for idx, img_b64 in enumerate(images_b64):
        try:
            encoded = img_b64.split(",", 1)[1] if "," in img_b64 else img_b64
            image = Image.open(io.BytesIO(base64.b64decode(encoded)))
            image = ImageOps.exif_transpose(image)

            processed_pil, processed_cv = preprocess_image(image)
            processed_images.append(base64.b64encode(cv2.imencode(".jpg", processed_cv)[1]).decode())

            best_img_data, current_raw = _best_rotation_ocr(processed_pil, idx)
            all_raw_text.append(current_raw)
            all_results.append(best_img_data)
        except Exception as e:
            logger.error(f"Error procesando imagen {idx}: {e}")
            continue

    for res in all_results:
        for k, v in res.items():
            if v:
                _merge_field(merged_fields, k, v)

    # Limpiar prefijos
    for k in ["fecha_recepcion", "fecha_rc"]:
        if merged_fields[k]:
            merged_fields[k] = merged_fields[k].replace("STAMP:", "").replace("NUMERIC:", "")

    return {"fields": merged_fields, "raw_text": "\n\n".join(all_raw_text), "processed_images": processed_images}


def generate_qr_base64(url: str) -> str:
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()
