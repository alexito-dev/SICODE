import io
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF

# Textos repetidos
_CENTER = "center"
_ENCONTRADO = "ENCONTRADO"
_HELVETICA = "Helvetica"
_LMARGIN = "LMARGIN"
_NO_ENCONTRADO = "NO ENCONTRADO"
_SHEET = "sheet"
_STATUS = "status"


_FONT_NAME = "Segoe UI"
_FONT_TITLE = Font(name=_FONT_NAME, size=14, bold=True, color="FFFFFF")
_FONT_SUB = Font(name=_FONT_NAME, size=9.5, italic=True, color="4B5563")
_FONT_KPI_VAL = Font(name=_FONT_NAME, size=11, color="111827", bold=True)
_FONT_TH = Font(name=_FONT_NAME, size=10, bold=True, color="FFFFFF")
_FONT_CELL = Font(name=_FONT_NAME, size=9.5)
_FONT_CELL_BOLD = Font(name=_FONT_NAME, size=9.5, bold=True)
_FONT_OK = Font(name=_FONT_NAME, size=9.5, bold=True, color="166534")
_FONT_MISS = Font(name=_FONT_NAME, size=9.5, bold=True, color="991B1B")
_FONT_WARN = Font(name=_FONT_NAME, size=9.5, bold=True, color="92400E")


def _solid_fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


_FILL_BANNER = _solid_fill("0F2942")
_FILL_TH = _solid_fill("0163AC")
_FILL_KPI = _solid_fill("F3F4F6")
_FILL_OK = _solid_fill("DCFCE7")
_FILL_MISS = _solid_fill("FEE2E2")
_FILL_WARN = _solid_fill("FEF3C7")
_FILL_ZEBRA = _solid_fill("F9FAFB")

_BORDER_SIDE = Side(style="thin", color="E5E7EB")
_BORDER_CELL = Border(left=_BORDER_SIDE, right=_BORDER_SIDE, top=_BORDER_SIDE, bottom=_BORDER_SIDE)
_ALIGN_CENTER = Alignment(horizontal=_CENTER, vertical=_CENTER)
_ALIGN_LEFT = Alignment(horizontal="left", vertical=_CENTER)


def _style_cell(cell, font, alignment, fill=None):
    cell.font = font
    cell.alignment = alignment
    cell.border = _BORDER_CELL
    if fill is not None:
        cell.fill = fill


def _write_headers(ws, row, headers):
    ws.row_dimensions[row].height = 26
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx, value=h)
        cell.font = _FONT_TH
        cell.fill = _FILL_TH
        cell.alignment = _ALIGN_CENTER
        cell.border = _BORDER_CELL


def _auto_width(ws):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


def _status_style(status):
    if status == _ENCONTRADO:
        return _FILL_OK, _FONT_OK
    if status == _NO_ENCONTRADO:
        return _FILL_MISS, _FONT_MISS
    return _FILL_WARN, _FONT_WARN


def _write_summary_banner(ws, user_name, results):
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total_rucs = len(results)
    found_count = sum(1 for r in results if r.get(_STATUS) == _ENCONTRADO)
    miss_count = sum(1 for r in results if r.get(_STATUS) == _NO_ENCONTRADO)

    ws.merge_cells("A1:F1")
    top_cell = ws["A1"]
    top_cell.value = "SICODE · Reporte Resumen de Clasificación ODI"
    top_cell.font = _FONT_TITLE
    top_cell.fill = _FILL_BANNER
    top_cell.alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:F2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Generado el {now_str} | Usuario: {user_name}"
    sub_cell.font = _FONT_SUB
    sub_cell.alignment = _ALIGN_CENTER
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[4].height = 24
    efectividad = f"{(found_count / total_rucs * 100):.1f}%" if total_rucs else "0%"
    kpis = [
        ("A4", "B4", "Total Consultados", total_rucs),
        ("C4", "C4", "Encontrados", found_count),
        ("D4", "D4", "No Encontrados", miss_count),
        ("E4", "F4", "% Efectividad", efectividad),
    ]
    for start_c, end_c, label, val in kpis:
        if start_c != end_c:
            ws.merge_cells(f"{start_c}:{end_c}")
        cell = ws[start_c]
        cell.value = f"{label}: {val}"
        _style_cell(cell, _FONT_KPI_VAL, _ALIGN_CENTER, _FILL_KPI)


def _write_summary_row(ws, i, res):
    row_idx = 6 + i
    ws.row_dimensions[row_idx].height = 20

    status = res.get(_STATUS, _NO_ENCONTRADO)
    files_str = ", ".join(res.get("files", [])) if res.get("files") else "—"
    sheets_str = ", ".join(res.get("sheets", [])) if res.get("sheets") else "—"
    st_fill, st_font = _status_style(status)
    row_fill = _FILL_ZEBRA if i % 2 == 0 else PatternFill(fill_type=None)

    columnas = [
        (i, _FONT_CELL, _ALIGN_CENTER, row_fill),
        (str(res.get("ruc", "")), _FONT_CELL_BOLD, _ALIGN_CENTER, row_fill),
        (status, st_font, _ALIGN_CENTER, st_fill),
        (res.get("matches_count", 0), _FONT_CELL, _ALIGN_CENTER, row_fill),
        (files_str, _FONT_CELL, _ALIGN_LEFT, row_fill),
        (sheets_str, _FONT_CELL, _ALIGN_LEFT, row_fill),
    ]
    for col_idx, (value, font, align, fill) in enumerate(columnas, 1):
        _style_cell(ws.cell(row=row_idx, column=col_idx, value=value), font, align, fill)


def _write_summary_sheet(ws, results, user_name):
    ws.title = "Resumen Clasificación"
    ws.views.sheetView[0].showGridLines = True
    _write_summary_banner(ws, user_name, results)
    _write_headers(
        ws, 6, ["N°", "RUC", "Estado", "N° Coincidencias", "Bases de Datos", "Hojas Coincidentes"]
    )
    for i, res in enumerate(results, 1):
        _write_summary_row(ws, i, res)

    _auto_width(ws)
    for letter, width in {"A": 8, "B": 16, "C": 18, "D": 16, "E": 30, "F": 35}.items():
        ws.column_dimensions[letter].width = width


def _group_matches_by_file(results):
    db_matches = {}
    for res in results:
        for match in res.get("matches", []):
            fname = f"{match.get('file_name', 'Base de Datos')}||{match.get('sheet', '—')}"
            db_matches.setdefault(fname, []).append(
                {"ruc": str(res.get("ruc", "")), _SHEET: match.get(_SHEET, "—"), "rows": match.get("rows", [])}
            )
    return db_matches


def _collect_columns(fmatches):
    f_cols = []
    f_seen = set()
    for item in fmatches:
        for r in item.get("rows", []):
            if not isinstance(r, dict):
                continue
            for k, v in r.items():
                k_s = str(k).strip()
                if k_s and v not in (None, "") and k_s not in f_seen:
                    f_seen.add(k_s)
                    f_cols.append(k_s)
    return f_cols


def _unique_tab_title(fname, seen_titles):
    """Nombre de pestaña limpio y único (máx 31 caracteres)."""
    file_part, sheet_part = fname.split("||", 1)
    base_name = file_part.replace(".xlsx", "").replace(".XLSX", "")
    raw_title = f"{base_name} - {sheet_part}"
    for ch in ("[", "]", ":", "*", "?", "/", "\\"):
        raw_title = raw_title.replace(ch, " ")
    tab_title = raw_title[:31]
    n = 1
    while tab_title.lower() in seen_titles:
        n += 1
        tab_title = raw_title[:28] + f"_{n}"
    seen_titles.add(tab_title.lower())
    return tab_title


def _write_match_row(ws, cur_r, seq, reg_idx, r_ruc, r_sheet, r_dict, f_cols):
    ws.row_dimensions[cur_r].height = 20
    row_fill = _FILL_ZEBRA if seq % 2 == 0 else PatternFill(fill_type=None)
    fijas = [
        (seq, _FONT_CELL, _ALIGN_CENTER),
        (r_ruc, _FONT_CELL_BOLD, _ALIGN_CENTER),
        (r_sheet, _FONT_CELL, _ALIGN_LEFT),
        (reg_idx, _FONT_CELL, _ALIGN_CENTER),
    ]
    for col_idx, (value, font, align) in enumerate(fijas, 1):
        _style_cell(ws.cell(row=cur_r, column=col_idx, value=value), font, align, row_fill)
    for c_off, col_key in enumerate(f_cols, 5):
        raw_val = r_dict.get(col_key, "") if isinstance(r_dict, dict) else ""
        val_str = str(raw_val) if raw_val is not None and raw_val != "" else ""
        _style_cell(ws.cell(row=cur_r, column=c_off, value=val_str), _FONT_CELL, _ALIGN_LEFT, row_fill)


def _write_file_sheet(ws, fmatches):
    ws.views.sheetView[0].showGridLines = True
    f_cols = _collect_columns(fmatches)
    _write_headers(ws, 1, ["N°", "RUC", "Hoja Coincidente", "Registro N°"] + f_cols)

    f_row_seq = 0
    seen_rows = set()
    for item in fmatches:
        r_ruc = item["ruc"]
        r_sheet = item[_SHEET]
        for reg_idx, r_dict in enumerate(item.get("rows", []), 1):
            row_sig = (r_ruc, r_sheet, json.dumps(r_dict, sort_keys=True, default=str))
            if row_sig in seen_rows:
                continue
            seen_rows.add(row_sig)
            f_row_seq += 1
            _write_match_row(ws, 1 + f_row_seq, f_row_seq, reg_idx, r_ruc, r_sheet, r_dict, f_cols)
    _auto_width(ws)


def generate_excel_export(results: list[dict], user_name: str = "Usuario") -> io.BytesIO:
    wb = openpyxl.Workbook()
    _write_summary_sheet(wb.active, results, user_name)

    seen_sheet_titles = {s.lower() for s in wb.sheetnames}
    for fname, fmatches in _group_matches_by_file(results).items():
        tab_title = _unique_tab_title(fname, seen_sheet_titles)
        _write_file_sheet(wb.create_sheet(title=tab_title), fmatches)

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def clean_pdf_text(val) -> str:
    if val is None:
        return ""
    s = str(val)
    replacements = {
        "—": "-",
        "–": "-",
        "•": "-",
        "·": "-",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "…": "...",
        "\u200b": "",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    return s.encode("latin-1", "replace").decode("latin-1")


class ODIPdfReport(FPDF):
    def __init__(self, user_name: str, now_str: str):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.user_name = clean_pdf_text(user_name)
        self.now_str = clean_pdf_text(now_str)
        self.set_auto_page_break(auto=True, margin=15)

    def header(self):
        self.set_text_color(0, 0, 0)
        self.set_font(_HELVETICA, "B", 11)
        self.set_xy(12, 6)
        self.cell(140, 5, "SICODE - SISTEMA DE CONTROL Y GESTION DOCUMENTAL", align="L")

        self.set_font(_HELVETICA, "", 8.5)
        self.set_xy(12, 11)
        self.cell(140, 4, "Modulo ODI - Reporte de Consulta y Clasificacion de RUCs", align="L")

        self.set_xy(150, 6)
        self.cell(135, 5, f"Fecha: {self.now_str}", align="R")
        self.set_xy(150, 11)
        self.cell(135, 4, f"Usuario: {self.user_name}", align="R")

        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.4)
        self.line(12, 17, 285, 17)
        self.set_y(22)

    def footer(self):
        self.set_y(-12)
        self.set_draw_color(160, 160, 160)
        self.set_line_width(0.3)
        self.line(12, self.get_y(), 285, self.get_y())
        self.set_font(_HELVETICA, "", 8)
        self.set_text_color(80, 80, 80)
        self.cell(140, 8, "SICODE - Documento de Consulta y Control", align="L")
        self.cell(133, 8, f"Pagina {self.page_no()} / {{nb}}", align="R")


def _pdf_title_bar(pdf, idx, ruc, status):
    """Barra de título de cada RUC (estilo impresión blanco y negro)."""
    pdf.set_fill_color(240, 240, 240)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.3)
    pdf.set_font(_HELVETICA, "B", 9.5)
    pdf.set_text_color(0, 0, 0)
    header_title = f"  #{idx}   |   RUC: {ruc}   -   ESTADO: {status}"
    pdf.cell(0, 6.5, header_title, border=1, fill=True, new_x=_LMARGIN, new_y="NEXT")


def _pdf_label_value(pdf, key, value, value_width, last):
    """Celda de campo (etiqueta) más celda de valor."""
    pdf.set_font(_HELVETICA, "B", 7.5)
    pdf.set_text_color(0, 0, 0)
    pdf.set_fill_color(248, 248, 248)
    pdf.cell(42, 5.5, f" {key}", border=1, fill=True)

    pdf.set_font(_HELVETICA, "", 8)
    pdf.set_text_color(0, 0, 0)
    if last:
        pdf.cell(value_width, 5.5, f" {value}", border=1, new_x=_LMARGIN, new_y="NEXT")
    else:
        pdf.cell(value_width, 5.5, f" {value}", border=1)


def _pdf_pair_row(pdf, pair1, pair2):
    pdf.set_draw_color(180, 180, 180)
    k1 = clean_pdf_text(str(pair1[0]))[:28]
    v1 = clean_pdf_text(str(pair1[1]))[:58]
    _pdf_label_value(pdf, k1, v1, 94, last=False)
    if pair2:
        k2 = clean_pdf_text(str(pair2[0]))[:28]
        v2 = clean_pdf_text(str(pair2[1]))[:58]
        _pdf_label_value(pdf, k2, v2, 95, last=True)
    else:
        pdf.cell(137, 5.5, "", border=1, new_x=_LMARGIN, new_y="NEXT")


def _pdf_record(pdf, reg_idx, r_dict, total_rows):
    if not isinstance(r_dict, dict) or not r_dict:
        return
    if total_rows > 1:
        if pdf.get_y() > 180:
            pdf.add_page()
        pdf.set_font(_HELVETICA, "I", 8)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(0, 4.5, f"      Registro #{reg_idx}:", border="LR", new_x=_LMARGIN, new_y="NEXT")

    items = [(k, v) for k, v in r_dict.items() if v not in (None, "")]
    # Cuadrícula horizontal de 2 columnas de Campo: Valor apta para impresión en A4 Landscape
    for i in range(0, len(items), 2):
        if pdf.get_y() > 185:
            pdf.add_page()
        pair2 = items[i + 1] if i + 1 < len(items) else None
        _pdf_pair_row(pdf, items[i], pair2)


def _pdf_match(pdf, match):
    fname = clean_pdf_text(str(match.get("file_name", "Base de Datos")))
    sheet = clean_pdf_text(str(match.get(_SHEET, "Hoja")))
    rows = match.get("rows", [])
    total_rows = len(rows)

    pdf.set_font(_HELVETICA, "B", 8.5)
    pdf.set_text_color(0, 0, 0)
    sub_info = f"      Base de Datos: {fname}   |   Hoja: {sheet}   ({total_rows} registro{'s' if total_rows > 1 else ''})"
    pdf.cell(0, 5.5, sub_info, border="LR", new_x=_LMARGIN, new_y="NEXT")

    for reg_idx, r_dict in enumerate(rows, 1):
        _pdf_record(pdf, reg_idx, r_dict, total_rows)


def _pdf_note(pdf, text):
    pdf.set_font(_HELVETICA, "I", 8.5)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 6, text, border="LRB", new_x=_LMARGIN, new_y="NEXT")


def _pdf_result_body(pdf, res, status):
    if status == _ENCONTRADO:
        for match in res.get("matches", []):
            _pdf_match(pdf, match)
    elif status == _NO_ENCONTRADO:
        _pdf_note(pdf, "      No se encontro informacion registrada para este RUC en las bases de datos activas.")
    else:
        _pdf_note(pdf, "      RUC invalido (debe contener exactamente 11 digitos numericos).")


def generate_pdf_export(results: list[dict], user_name: str = "Usuario") -> bytes:
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    pdf = ODIPdfReport(user_name=user_name, now_str=now_str)
    pdf.alias_nb_pages()
    pdf.add_page()

    for idx, res in enumerate(results, 1):
        if pdf.get_y() > 175:
            pdf.add_page()

        ruc = clean_pdf_text(str(res.get("ruc", "")))
        status = clean_pdf_text(str(res.get(_STATUS, _NO_ENCONTRADO)))

        _pdf_title_bar(pdf, idx, ruc, status)
        _pdf_result_body(pdf, res, status)
        pdf.ln(3)

    return bytes(pdf.output())
