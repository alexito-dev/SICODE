import io
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF


def generate_excel_export(results: list[dict], user_name: str = "Usuario") -> io.BytesIO:
    wb = openpyxl.Workbook()

    # Estilos compartidos
    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=9.5, italic=True, color="4B5563")
    font_kpi_val = Font(name="Segoe UI", size=11, color="111827", bold=True)
    font_th = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_cell = Font(name="Segoe UI", size=9.5)
    font_cell_bold = Font(name="Segoe UI", size=9.5, bold=True)

    fill_banner = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
    fill_th = PatternFill(start_color="0163AC", end_color="0163AC", fill_type="solid")
    PatternFill(start_color="107C41", end_color="107C41", fill_type="solid")
    fill_kpi = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")

    fill_ok = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_miss = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_warn = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    font_ok = Font(name="Segoe UI", size=9.5, bold=True, color="166534")
    font_miss = Font(name="Segoe UI", size=9.5, bold=True, color="991B1B")
    font_warn = Font(name="Segoe UI", size=9.5, bold=True, color="92400E")

    thin_border_side = Side(style="thin", color="E5E7EB")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_th = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total_rucs = len(results)
    found_count = sum(1 for r in results if r.get("status") == "ENCONTRADO")
    miss_count = sum(1 for r in results if r.get("status") == "NO ENCONTRADO")

    # ── Hoja 2: Resumen Clasificación ─────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Resumen Clasificación"
    ws_sum.views.sheetView[0].showGridLines = True

    # Banner de título
    ws_sum.merge_cells("A1:F1")
    top_cell = ws_sum["A1"]
    top_cell.value = "SICODE · Reporte Resumen de Clasificación ODI"
    top_cell.font = font_title
    top_cell.fill = fill_banner
    top_cell.alignment = align_center
    ws_sum.row_dimensions[1].height = 36

    # Subtítulo con metadata
    ws_sum.merge_cells("A2:F2")
    sub_cell = ws_sum["A2"]
    sub_cell.value = f"Generado el {now_str} | Usuario: {user_name}"
    sub_cell.font = font_sub
    sub_cell.alignment = align_center
    ws_sum.row_dimensions[2].height = 20

    # Tarjetas KPI (Fila 4)
    ws_sum.row_dimensions[4].height = 24
    kpis = [
        ("A4", "B4", "Total Consultados", total_rucs),
        ("C4", "C4", "Encontrados", found_count),
        ("D4", "D4", "No Encontrados", miss_count),
        ("E4", "F4", "% Efectividad", f"{(found_count / total_rucs * 100):.1f}%" if total_rucs else "0%"),
    ]
    for start_c, end_c, label, val in kpis:
        if start_c != end_c:
            ws_sum.merge_cells(f"{start_c}:{end_c}")
        cell = ws_sum[start_c]
        cell.value = f"{label}: {val}"
        cell.font = font_kpi_val
        cell.fill = fill_kpi
        cell.alignment = align_center
        cell.border = border_cell

    # Encabezados de tabla (Fila 6)
    headers = ["N°", "RUC", "Estado", "N° Coincidencias", "Bases de Datos", "Hojas Coincidentes"]
    ws_sum.row_dimensions[6].height = 26
    for col_idx, h in enumerate(headers, 1):
        cell = ws_sum.cell(row=6, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = align_center
        cell.border = border_th

    # Filas de datos
    for i, res in enumerate(results, 1):
        row_idx = 6 + i
        ws_sum.row_dimensions[row_idx].height = 20

        status = res.get("status", "NO ENCONTRADO")
        files_str = ", ".join(res.get("files", [])) if res.get("files") else "—"
        sheets_str = ", ".join(res.get("sheets", [])) if res.get("sheets") else "—"
        matches_count = res.get("matches_count", 0)

        if status == "ENCONTRADO":
            st_fill = fill_ok
            st_font = font_ok
        elif status == "NO ENCONTRADO":
            st_fill = fill_miss
            st_font = font_miss
        else:
            st_fill = fill_warn
            st_font = font_warn

        row_fill = fill_zebra if i % 2 == 0 else PatternFill(fill_type=None)

        # N°
        c1 = ws_sum.cell(row=row_idx, column=1, value=i)
        c1.alignment = align_center
        c1.font = font_cell
        c1.border = border_cell
        c1.fill = row_fill

        # RUC
        c2 = ws_sum.cell(row=row_idx, column=2, value=str(res.get("ruc", "")))
        c2.alignment = align_center
        c2.font = font_cell_bold
        c2.border = border_cell
        c2.fill = row_fill

        # Estado
        c3 = ws_sum.cell(row=row_idx, column=3, value=status)
        c3.alignment = align_center
        c3.font = st_font
        c3.fill = st_fill
        c3.border = border_cell

        # N° Coincidencias
        c4 = ws_sum.cell(row=row_idx, column=4, value=matches_count)
        c4.alignment = align_center
        c4.font = font_cell
        c4.border = border_cell
        c4.fill = row_fill

        # Bases de Datos
        c5 = ws_sum.cell(row=row_idx, column=5, value=files_str)
        c5.alignment = align_left
        c5.font = font_cell
        c5.border = border_cell
        c5.fill = row_fill

        # Hojas Coincidentes
        c6 = ws_sum.cell(row=row_idx, column=6, value=sheets_str)
        c6.alignment = align_left
        c6.font = font_cell
        c6.border = border_cell
        c6.fill = row_fill

    # Auto-ancho columnas Hoja 2
    for col in ws_sum.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 12)
    ws_sum.column_dimensions["A"].width = 8
    ws_sum.column_dimensions["B"].width = 16
    ws_sum.column_dimensions["C"].width = 18
    ws_sum.column_dimensions["D"].width = 16
    ws_sum.column_dimensions["E"].width = 30
    ws_sum.column_dimensions["F"].width = 35

    # ── Hojas Opcionales Específicas por Base de Datos (si hay coincidencias) ──
    db_matches = {}
    for res in results:
        for match in res.get("matches", []):
            fname = f"{match.get('file_name', 'Base de Datos')}||{match.get('sheet', '—')}"
            db_matches.setdefault(fname, []).append(
                {"ruc": str(res.get("ruc", "")), "sheet": match.get("sheet", "—"), "rows": match.get("rows", [])}
            )

    seen_sheet_titles = {s.lower() for s in wb.sheetnames}
    for fname, fmatches in db_matches.items():
        f_cols = []
        f_seen = set()
        for item in fmatches:
            for r in item.get("rows", []):
                if isinstance(r, dict):
                    for k, v in r.items():
                        k_s = str(k).strip()
                        if k_s and v not in (None, "") and k_s not in f_seen:
                            f_seen.add(k_s)
                            f_cols.append(k_s)

        # Nombre de pestaña limpio y único (máx 31 caracteres)
        file_part, sheet_part = fname.split("||", 1)
        base_name = file_part.replace(".xlsx", "").replace(".XLSX", "")
        raw_title = f"{base_name} - {sheet_part}"
        for ch in ("[", "]", ":", "*", "?", "/", "\\"):
            raw_title = raw_title.replace(ch, " ")
        tab_title = raw_title[:31]
        n = 1
        while tab_title.lower() in seen_sheet_titles:
            n += 1
            tab_title = raw_title[:28] + f"_{n}"
        seen_sheet_titles.add(tab_title.lower())

        ws_file = wb.create_sheet(title=tab_title)
        ws_file.views.sheetView[0].showGridLines = True

        f_headers = ["N°", "RUC", "Hoja Coincidente", "Registro N°"] + f_cols
        ws_file.row_dimensions[1].height = 26
        for c_idx, h in enumerate(f_headers, 1):
            cell = ws_file.cell(row=1, column=c_idx, value=h)
            cell.font = font_th
            cell.fill = fill_th
            cell.alignment = align_center
            cell.border = border_th

        f_row_seq = 0
        seen_rows = set()
        for item in fmatches:
            r_ruc = item["ruc"]
            r_sheet = item["sheet"]
            for reg_idx, r_dict in enumerate(item.get("rows", []), 1):
                row_sig = (r_ruc, r_sheet, json.dumps(r_dict, sort_keys=True, default=str))
                if row_sig in seen_rows:
                    continue
                seen_rows.add(row_sig)
                f_row_seq += 1
                cur_r = 1 + f_row_seq
                ws_file.row_dimensions[cur_r].height = 20
                row_fill = fill_zebra if f_row_seq % 2 == 0 else PatternFill(fill_type=None)

                c1 = ws_file.cell(row=cur_r, column=1, value=f_row_seq)
                c1.alignment = align_center
                c1.font = font_cell
                c1.border = border_cell
                c1.fill = row_fill

                c2 = ws_file.cell(row=cur_r, column=2, value=r_ruc)
                c2.alignment = align_center
                c2.font = font_cell_bold
                c2.border = border_cell
                c2.fill = row_fill

                c3 = ws_file.cell(row=cur_r, column=3, value=r_sheet)
                c3.alignment = align_left
                c3.font = font_cell
                c3.border = border_cell
                c3.fill = row_fill

                c4 = ws_file.cell(row=cur_r, column=4, value=reg_idx)
                c4.alignment = align_center
                c4.font = font_cell
                c4.border = border_cell
                c4.fill = row_fill

                for c_off, col_key in enumerate(f_cols, 5):
                    raw_val = r_dict.get(col_key, "") if isinstance(r_dict, dict) else ""
                    val_str = str(raw_val) if raw_val is not None and raw_val != "" else ""
                    c_val = ws_file.cell(row=cur_r, column=c_off, value=val_str)
                    c_val.font = font_cell
                    c_val.alignment = align_left
                    c_val.border = border_cell
                    c_val.fill = row_fill

        for col in ws_file.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_file.column_dimensions[col_letter].width = max(max_len + 4, 12)

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
        self.set_font("Helvetica", "B", 11)
        self.set_xy(12, 6)
        self.cell(140, 5, "SICODE - SISTEMA DE CONTROL Y GESTION DOCUMENTAL", align="L")

        self.set_font("Helvetica", "", 8.5)
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
        self.set_font("Helvetica", "", 8)
        self.set_text_color(80, 80, 80)
        self.cell(140, 8, "SICODE - Documento de Consulta y Control", align="L")
        self.cell(133, 8, f"Pagina {self.page_no()} / {{nb}}", align="R")


def generate_pdf_export(results: list[dict], user_name: str = "Usuario") -> bytes:
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    pdf = ODIPdfReport(user_name=user_name, now_str=now_str)
    pdf.alias_nb_pages()
    pdf.add_page()

    for idx, res in enumerate(results, 1):
        if pdf.get_y() > 175:
            pdf.add_page()

        ruc = clean_pdf_text(str(res.get("ruc", "")))
        status = clean_pdf_text(str(res.get("status", "NO ENCONTRADO")))

        # Barra de título de cada RUC (Estilo impresión Blanco y Negro)
        pdf.set_fill_color(240, 240, 240)
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(0, 0, 0)
        header_title = f"  #{idx}   |   RUC: {ruc}   -   ESTADO: {status}"
        pdf.cell(0, 6.5, header_title, border=1, fill=True, new_x="LMARGIN", new_y="NEXT")

        if status == "ENCONTRADO":
            matches = res.get("matches", [])
            for match in matches:
                fname = clean_pdf_text(str(match.get("file_name", "Base de Datos")))
                sheet = clean_pdf_text(str(match.get("sheet", "Hoja")))
                rows = match.get("rows", [])
                total_rows = len(rows)

                # Subtítulo con base de datos y hoja
                pdf.set_font("Helvetica", "B", 8.5)
                pdf.set_text_color(0, 0, 0)
                sub_info = f"      Base de Datos: {fname}   |   Hoja: {sheet}   ({total_rows} registro{'s' if total_rows > 1 else ''})"
                pdf.cell(0, 5.5, sub_info, border="LR", new_x="LMARGIN", new_y="NEXT")

                for reg_idx, r_dict in enumerate(rows, 1):
                    if not isinstance(r_dict, dict) or not r_dict:
                        continue

                    if total_rows > 1:
                        if pdf.get_y() > 180:
                            pdf.add_page()
                        pdf.set_font("Helvetica", "I", 8)
                        pdf.set_text_color(60, 60, 60)
                        pdf.cell(0, 4.5, f"      Registro #{reg_idx}:", border="LR", new_x="LMARGIN", new_y="NEXT")

                    items = [(k, v) for k, v in r_dict.items() if v not in (None, "")]
                    if not items:
                        continue
                    # Cuadrícula horizontal de 2 columnas de Campo: Valor apta para impresión en A4 Landscape
                    for i in range(0, len(items), 2):
                        if pdf.get_y() > 185:
                            pdf.add_page()

                        pair1 = items[i]
                        pair2 = items[i + 1] if i + 1 < len(items) else None

                        k1 = clean_pdf_text(str(pair1[0]))[:28]
                        v1 = clean_pdf_text(str(pair1[1]))[:58]

                        pdf.set_draw_color(180, 180, 180)

                        # Columna 1
                        pdf.set_font("Helvetica", "B", 7.5)
                        pdf.set_text_color(0, 0, 0)
                        pdf.set_fill_color(248, 248, 248)
                        pdf.cell(42, 5.5, f" {k1}", border=1, fill=True)

                        pdf.set_font("Helvetica", "", 8)
                        pdf.set_text_color(0, 0, 0)
                        pdf.cell(94, 5.5, f" {v1}", border=1)

                        # Columna 2
                        if pair2:
                            k2 = clean_pdf_text(str(pair2[0]))[:28]
                            v2 = clean_pdf_text(str(pair2[1]))[:58]

                            pdf.set_font("Helvetica", "B", 7.5)
                            pdf.set_text_color(0, 0, 0)
                            pdf.set_fill_color(248, 248, 248)
                            pdf.cell(42, 5.5, f" {k2}", border=1, fill=True)

                            pdf.set_font("Helvetica", "", 8)
                            pdf.set_text_color(0, 0, 0)
                            pdf.cell(95, 5.5, f" {v2}", border=1, new_x="LMARGIN", new_y="NEXT")
                        else:
                            pdf.cell(137, 5.5, "", border=1, new_x="LMARGIN", new_y="NEXT")
        elif status == "NO ENCONTRADO":
            pdf.set_font("Helvetica", "I", 8.5)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(
                0,
                6,
                "      No se encontro informacion registrada para este RUC en las bases de datos activas.",
                border="LRB",
                new_x="LMARGIN",
                new_y="NEXT",
            )
        else:
            pdf.set_font("Helvetica", "I", 8.5)
            pdf.set_text_color(80, 80, 80)
            pdf.cell(
                0,
                6,
                "      RUC invalido (debe contener exactamente 11 digitos numericos).",
                border="LRB",
                new_x="LMARGIN",
                new_y="NEXT",
            )

        pdf.ln(3)

    return bytes(pdf.output())
