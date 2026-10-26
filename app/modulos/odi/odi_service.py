"""
Módulo ODI - Clasificador de RUCs por hoja de Excel.
Lee un Excel en memoria (sin guardarlo en el servidor) y determina,
para cada RUC de 11 dígitos encontrado, en qué hoja(s) aparece y
con qué datos de fila (columna -> valor) figura en cada una.
"""

import io
import re
import openpyxl


_STOP_WORDS = {
    "NOMBRES",
    "APELLIDOS",
    "NOMBRE",
    "FECHA",
    "ESTADO",
    "CÓDIGO",
    "CODIGO",
    "TOTAL",
    "PROYECTO",
    "HOJA",
    "RESUMEN",
    "CONSOLIDADO",
    "SI",
    "NO",
    "ACTIVO",
    "INACTIVO",
    "COLUMNA",
}


def _normalize_key_cell(val) -> str | None:
    """Texto limpio de la celda, o None si no aporta (vacía, fecha ISO o decimal no entero)."""
    if val is None:
        return None
    s = str(val).strip().replace("\xa0", " ")
    if not s or re.search(r"\d{4}-\d{2}-\d{2}", s):
        return None

    # Tratar floats que son enteros
    if "." in s or "e" in s.lower():
        try:
            f = float(s)
        except Exception:
            return s
        return str(int(f)) if f.is_integer() else None
    return s


def _classify_numeric_key(s_clean: str, rucs: set, dnis: set, codes: set) -> None:
    if len(s_clean) == 11:
        rucs.add(s_clean)
    elif len(s_clean) == 8:
        dnis.add(s_clean)
    elif 7 <= len(s_clean) <= 12:
        codes.add(s_clean)


def _classify_text_key(s: str, rucs: set, dnis: set, names: set) -> None:
    for m in re.finditer(r"\b([12]\d{10})\b", s):
        rucs.add(m.group(1))
    for m in re.finditer(r"\b(\d{8})\b", s):
        dnis.add(m.group(1))

    clean_text = " ".join(s.upper().split())
    if 3 <= len(clean_text) <= 80 and not re.search(r"^\d+$", clean_text):
        if clean_text not in _STOP_WORDS and len(clean_text) >= 4:
            names.add(clean_text)


def _extract_keys_from_row(row) -> list[str]:
    rucs, dnis, codes, names = set(), set(), set(), set()

    for val in row:
        s = _normalize_key_cell(val)
        if s is None:
            continue
        s_clean = s.replace("-", "").replace(" ", "")
        if s_clean.isdigit():
            _classify_numeric_key(s_clean, rucs, dnis, codes)
        else:
            _classify_text_key(s, rucs, dnis, names)

    return list(rucs | dnis | codes | names)


def _format_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _unique_headers(header_row) -> list[str]:
    """Nombres de columna únicos; las celdas vacías pasan a 'Columna N'."""
    headers = []
    seen_h = set()
    for i, h in enumerate(header_row):
        h_name = str(h).strip() if h not in (None, "") else f"Columna {i + 1}"
        orig_name = h_name
        dup_idx = 2
        while h_name in seen_h:
            h_name = f"{orig_name}_{dup_idx}"
            dup_idx += 1
        seen_h.add(h_name)
        headers.append(h_name)
    return headers


def _first_row_is_data(first_row) -> bool:
    """La primera fila es de datos si ya trae un RUC o DNI numérico claro."""
    first_keys = _extract_keys_from_row(first_row)
    return bool(first_keys) and any(k.isdigit() for k in first_keys)


def _row_to_dict(row, headers) -> dict:
    row_data = {}
    for i, label in enumerate(headers):
        val = row[i] if i < len(row) else None
        row_data[label] = _format_cell(val) if val is not None and val != "" else ""
    for i in range(len(headers), len(row)):
        val = row[i]
        if val is not None and val != "":
            row_data[f"Columna {i + 1}"] = _format_cell(val)
    return row_data


def _analyze_sheet(ws, ruc_map: dict) -> dict:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {"name": ws.title, "ruc_count": 0, "headers": []}

    headers = _unique_headers(rows[0])
    start_idx = 1
    if _first_row_is_data(rows[0]):
        start_idx = 0
        headers = [f"Columna {i + 1}" for i in range(len(rows[0]))]

    count = 0
    sheet_ruc_rows: dict[str, list[dict]] = {}
    for row in rows[start_idx:]:
        keys = _extract_keys_from_row(row)
        if not keys:
            continue
        count += 1
        row_data = _row_to_dict(row, headers)
        for key in keys:
            sheet_ruc_rows.setdefault(key, []).append(row_data)

    for key, r_list in sheet_ruc_rows.items():
        ruc_map.setdefault(key, []).append({"sheet": ws.title, "rows": r_list})
    return {"name": ws.title, "ruc_count": count, "headers": headers}


def analyze_excel(content: bytes) -> dict:
    """
    Recorre todas las hojas del Excel y arma un mapa Clave/RUC/DNI/Nombre -> [{sheet, rows}],
    capturando la información completa de la fila.
    """
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ruc_map: dict[str, list[dict]] = {}
    sheets = [_analyze_sheet(ws, ruc_map) for ws in wb.worksheets]
    wb.close()
    return {
        "sheets": sheets,
        "ruc_map": ruc_map,
        "total_rucs": len(ruc_map),
    }
