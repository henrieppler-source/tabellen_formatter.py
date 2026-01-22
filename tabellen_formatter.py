import os
import glob
import re
from copy import copy as copy_style
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, PatternFill


# ------------------------------------------------------------
# Konfiguration
# ------------------------------------------------------------

LAYOUT_DIR = "Layouts"
OUTPUT_DIR = "Ausgabedateien"

RAW_SHEET_NAMES = {
    1: "XML-Tab1-Land",
    2: "XML-Tab2-Land",
    3: "XML-Tab3-Land",
    5: "XML-Tab5-Land",
}

TEMPLATES = {
    1: {"ext": "Tabelle-1-Layout_g.xlsx", "int": "Tabelle-1-Layout_INTERN.xlsx"},
    2: {"ext": "Tabelle-2-Layout_g.xlsx", "int": "Tabelle-2-Layout_INTERN.xlsx"},
    3: {"ext": "Tabelle-3-Layout_g.xlsx", "int": "Tabelle-3-Layout_INTERN.xlsx"},
    5: {"ext": "Tabelle-5-Layout_g.xlsx", "int": "Tabelle-5-Layout_INTERN.xlsx"},
}


# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

def is_numeric_like(v):
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip().replace(".", "").replace(",", "")
        return s.isdigit() or v.strip() in ["-", "X"]
    return False


def extract_month_from_raw(ws, table_no):
    if table_no == 1:
        return ws.cell(row=3, column=1).value
    elif table_no in (2, 3):
        return ws.cell(row=4, column=1).value
    elif table_no == 5:
        return ws.cell(row=3, column=1).value
    return None


def extract_stand_from_raw(ws, max_search_rows=40):
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(max_row, max(max_row - max_search_rows, 1) - 1, -1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "Stand:" in v:
                return v.strip()
    return None


def update_footer_with_stand_and_copyright(ws, stand_text):
    max_row = ws.max_row
    max_col = ws.max_column
    current_year = datetime.now().year

    copyright_row = None
    for r in range(max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            text = v

            def repl(_m):
                return f"(C)opyright {current_year}"

            new_text = re.sub(r"\(C\)opyright\s+\d{4}", repl, text)
            ws.cell(row=r, column=1).value = new_text
            copyright_row = r
            break

    if not copyright_row or not stand_text:
        return

    stand_col = None
    for c in range(1, max_col + 1):
        v = ws.cell(row=copyright_row, column=c).value
        if isinstance(v, str) and "Stand:" in v:
            stand_col = c
            break

    # Entferne andere "Stand:"-Vorkommen (außer in Copyright-Zeile)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().startswith("Stand:") and r != copyright_row:
                ws.cell(row=r, column=c).value = ""

    if stand_col is None:
        stand_col = max_col  # Fallback

    cop_cell = ws.cell(row=copyright_row, column=1)
    tgt = ws.cell(row=copyright_row, column=stand_col)
    tgt.value = stand_text

    # Stil wie Copyright
    tgt.font = copy_style(cop_cell.font)
    tgt.border = copy_style(cop_cell.border)
    tgt.fill = copy_style(cop_cell.fill)
    tgt.number_format = cop_cell.number_format
    tgt.protection = copy_style(cop_cell.protection)
    tgt.alignment = Alignment(
        horizontal="right",
        vertical=cop_cell.alignment.vertical if cop_cell.alignment else "center",
    )


def get_merged_secondary_checker(ws):
    merged = list(ws.merged_cells.ranges)

    def is_secondary(row, col):
        for rg in merged:
            if rg.min_row <= row <= rg.max_row and rg.min_col <= col <= rg.max_col:
                return not (row == rg.min_row and col == rg.min_col)
        return False

    return is_secondary


def mark_cells_with_1_or_2(ws, col_index, fill):
    max_row = ws.max_row
    for r in range(1, max_row + 1):
        cell = ws.cell(row=r, column=col_index)
        v = cell.value
        if isinstance(v, (int, float)) and v in (1, 2):
            cell.fill = fill
        elif isinstance(v, str) and v.strip() in ("1", "2"):
            cell.fill = fill


def format_numeric_cells(ws, skip_cols=None):
    """
    Ganzzahlen mit festem Leerzeichen als Tausendertrennzeichen
    (auch bei Millionen+), ohne Dezimalstellen.

    Regeln:
    - 0 bleibt 0
    - Zellen mit "-" (Text) werden ignoriert
    - Zellen mit "X" (Text) werden ignoriert
    - Prozent-/Kommaspalten werden über skip_cols ausgeschlossen
    - Negative Zahlen: "- " + Zahl (Minus + genau ein Leerzeichen)
    """
    if skip_cols is None:
        skip_cols = set()

    # Gruppierung mit fixem Leerzeichen bis sehr große Zahlen:
    # z.B. 17 982 291
    pos = "#\\ ###\\ ###\\ ###\\ ###\\ ##0"
    neg = "-\\ " + pos
    thousands_format = f"{pos};{neg};0"

    for row in ws.iter_rows():
        for cell in row:
            if cell.column in skip_cols:
                continue

            v = cell.value
            if v in ("-", "X"):
                continue

            if isinstance(v, (int, float)):
                if isinstance(v, float):
                    cell.value = int(round(v))
                cell.number_format = thousands_format


# ------------------------------------------------------------
# Tabelle 1
# ------------------------------------------------------------

def build_table1_workbook(raw_path, template_path, internal_layout):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    month_text = extract_month_from_raw(ws_raw, 1)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]

    if internal_layout:
        ws.cell(row=5, column=1).value = month_text
    else:
        ws.cell(row=3, column=1).value = month_text

    is_sec = get_merged_secondary_checker(ws)
    max_col_ws = ws.max_column

    def detect_data_and_footer(sheet, numeric_col=4):
        max_row = sheet.max_row
        first_data = None
        for r in range(1, max_row + 1):
            if is_numeric_like(sheet.cell(row=r, column=numeric_col).value):
                first_data = r
                break
        if first_data is None:
            first_data = 1

        footnote_start = max_row + 1
        for r in range(1, max_row + 1):
            v = sheet.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip().startswith("-"):
                footnote_start = r
                break
        return first_data, footnote_start

    fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=4)
    fdr_t, ft_t = detect_data_and_footer(ws, numeric_col=4)

    n_rows = min(ft_raw - fdr_raw, ft_t - fdr_t)

    for offset in range(n_rows):
        r_raw = fdr_raw + offset
        r_t = fdr_t + offset
        for c in range(1, max_col_ws + 1):
            if is_sec(r_t, c):
                continue
            ws.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws, stand_text)

    # Tabelle 1: Spalte I (9) ist Kommazahl/„Prozent“ -> NICHT formatieren
    format_numeric_cells(ws, skip_cols={9})

    return wb


def process_table1(raw_path, tmpl_ext_path, tmpl_int_path, is_jj):
    print(f"Verarbeite Tabelle 1 aus '{raw_path}' ...")

    wb_int = None
    if os.path.exists(tmpl_int_path):
        wb_int = build_table1_workbook(raw_path, tmpl_int_path, internal_layout=True)
        base = os.path.splitext(os.path.basename(raw_path))[0]
        out_int = os.path.join(OUTPUT_DIR, base + "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern nicht gefunden: {tmpl_int_path}")

    if is_jj:
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table1_workbook(raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            ws_ext = wb_ext[wb_ext.sheetnames[0]]
            ws_ext.cell(row=1, column=1).value = None

            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            mark_cells_with_1_or_2(ws_ext, 7, fill)  # Tabelle 1: Spalte G

            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern (JJ): {out_ext}")
    else:
        if os.path.exists(tmpl_ext_path):
            wb_ext = build_table1_workbook(raw_path, tmpl_ext_path, internal_layout=False)
            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern: {out_ext}")
        else:
            print(f"  [WARNUNG] Vorlage extern nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Tabelle 2 & 3
# ------------------------------------------------------------

def build_table2_or_3_workbook(table_no, raw_path, template_path, internal_layout):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]]

    month_text = extract_month_from_raw(ws_raw, table_no)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]

    if internal_layout:
        ws.cell(row=6, column=1).value = month_text
    else:
        ws.cell(row=3, column=1).value = month_text

    is_sec = get_merged_secondary_checker(ws)
    max_col_t = ws.max_column

    def detect_data_and_footer(sheet, numeric_col=3):
        max_row = sheet.max_row
        first_data = None
        for r in range(1, max_row + 1):
            if is_numeric_like(sheet.cell(row=r, column=numeric_col).value):
                first_data = r
                break
        if first_data is None:
            first_data = 1

        footnote_start = max_row + 1
        for r in range(1, max_row + 1):
            v = sheet.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip().startswith("-"):
                footnote_start = r
                break
        return first_data, footnote_start

    fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=3)
    fdr_t, ft_t = detect_data_and_footer(ws, numeric_col=3)

    n_rows = min(ft_raw - fdr_raw, ft_t - fdr_t)

    for offset in range(n_rows):
        r_raw = fdr_raw + offset
        r_t = fdr_t + offset
        for c in range(3, max_col_t + 1):
            if is_sec(r_t, c):
                continue
            ws.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws, stand_text)

    # Tabelle 2 & 3: Spalte G (7) ist Kommazahl/„Prozent“ -> NICHT formatieren
    format_numeric_cells(ws, skip_cols={7})

    return wb


def process_table2_or_3(table_no, raw_path, tmpl_ext_path, tmpl_int_path, is_jj):
    print(f"Verarbeite Tabelle {table_no} aus '{raw_path}' ...")

    wb_int = None
    if os.path.exists(tmpl_int_path):
        wb_int = build_table2_or_3_workbook(table_no, raw_path, tmpl_int_path, internal_layout=True)
        base = os.path.splitext(os.path.basename(raw_path))[0]
        out_int = os.path.join(OUTPUT_DIR, base + "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern nicht gefunden: {tmpl_int_path}")

    if is_jj:
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table2_or_3_workbook(table_no, raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            ws_ext = wb_ext[wb_ext.sheetnames[0]]
            ws_ext.cell(row=1, column=1).value = None

            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            mark_cells_with_1_or_2(ws_ext, 5, fill)  # Tabelle 2/3: Spalte E

            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern (JJ): {out_ext}")
    else:
        if os.path.exists(tmpl_ext_path):
            wb_ext = build_table2_or_3_workbook(table_no, raw_path, tmpl_ext_path, internal_layout=False)
            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern: {out_ext}")
        else:
            print(f"  [WARNUNG] Vorlage extern nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Tabelle 5 (5 Blätter)
# ------------------------------------------------------------

def build_table5_workbook(raw_path, template_path, internal_layout):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[5]]

    month_text = extract_month_from_raw(ws_raw, 5)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    max_row = ws_raw.max_row

    starts = []
    for r in range(1, max_row + 1):
        v = ws_raw.cell(row=r, column=2).value
        if isinstance(v, str) and re.match(r"Bayern\s+\d\)", v.strip()):
            starts.append(r)

    block_ranges = []
    for i, start in enumerate(starts):
        end = (starts[i + 1] - 1) if i < len(starts) - 1 else max_row

        last_nonempty = start
        for rr in range(start, end + 1):
            if any(ws_raw.cell(row=rr, column=c).value not in (None, "") for c in range(1, 10)):
                last_nonempty = rr
        block_ranges.append((start, last_nonempty))

    def fill_sheet_from_block(ws_t, start_row, end_row):
        is_sec = get_merged_secondary_checker(ws_t)
        max_row_t = ws_t.max_row

        first_data_t = None
        for r in range(1, max_row_t + 1):
            if is_numeric_like(ws_t.cell(row=r, column=3).value):
                first_data_t = r
                break
        if first_data_t is None:
            return

        raw_r = start_row
        t_r = first_data_t
        while raw_r <= end_row and t_r <= max_row_t:
            for c in range(3, 9):  # C..H
                if is_sec(t_r, c):
                    continue
                ws_t.cell(row=t_r, column=c).value = ws_raw.cell(row=raw_r, column=c).value
            raw_r += 1
            t_r += 1

    for i, (start, end) in enumerate(block_ranges):
        if i >= len(wb.worksheets):
            break
        ws_t = wb.worksheets[i]

        if internal_layout:
            ws_t.cell(row=5, column=1).value = month_text
        else:
            ws_t.cell(row=3, column=1).value = month_text

        fill_sheet_from_block(ws_t, start, end)
        update_footer_with_stand_and_copyright(ws_t, stand_text)

        # Tabelle 5: Spalte H (8) ist Kommazahl/„Prozent“ -> NICHT formatieren
        format_numeric_cells(ws_t, skip_cols={8})

    return wb


def process_table5(raw_path, tmpl_ext_path, tmpl_int_path, is_jj):
    print(f"Verarbeite Tabelle 5 aus '{raw_path}' ...")

    wb_int = None
    if os.path.exists(tmpl_int_path):
        wb_int = build_table5_workbook(raw_path, tmpl_int_path, internal_layout=True)
        base = os.path.splitext(os.path.basename(raw_path))[0]
        out_int = os.path.join(OUTPUT_DIR, base + "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern nicht gefunden: {tmpl_int_path}")

    if is_jj:
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table5_workbook(raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

            for ws_ext in wb_ext.worksheets:
                ws_ext.cell(row=1, column=1).value = None
                mark_cells_with_1_or_2(ws_ext, 6, fill)  # Tabelle 5 JJ: Spalte F

            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern (JJ): {out_ext}")
    else:
        if os.path.exists(tmpl_ext_path):
            wb_ext = build_table5_workbook(raw_path, tmpl_ext_path, internal_layout=False)
            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern: {out_ext}")
        else:
            print(f"  [WARNUNG] Vorlage extern nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Sammelmappen erzeugen
# ------------------------------------------------------------

def detect_period_from_filename(filename: str) -> str | None:
    """
    Erkennt Periode aus Dateinamen:
      - YYYY-MM
      - YYYY-Q[1-4]
      - YYYY-H[1-2]
      - YYYY-JJ
    """
    name = os.path.basename(filename)

    m = re.search(r"(20\d{2}-(?:0[1-9]|1[0-2]))", name)
    if m:
        return m.group(1)

    m = re.search(r"(20\d{2}-Q[1-4])", name)
    if m:
        return m.group(1)

    m = re.search(r"(20\d{2}-H[12])", name)
    if m:
        return m.group(1)

    m = re.search(r"(20\d{2}-JJ)", name)
    if m:
        return m.group(1)

    return None


def copy_sheet_to_workbook(src_ws, tgt_wb, new_title: str):
    """
    Kopiert ein komplettes Worksheet inkl. Werte, Styles, Dimensionen, Merges.
    Damit Layout 1:1 bleibt.
    """
    tgt_ws = tgt_wb.create_sheet(title=new_title)

    # Sheet properties / view / page setup
    tgt_ws.sheet_format = copy_style(src_ws.sheet_format)
    tgt_ws.sheet_properties = copy_style(src_ws.sheet_properties)
    tgt_ws.sheet_view = copy_style(src_ws.sheet_view)
    tgt_ws.page_setup = copy_style(src_ws.page_setup)
    tgt_ws.page_margins = copy_style(src_ws.page_margins)
    tgt_ws.print_options = copy_style(src_ws.print_options)
    tgt_ws.protection = copy_style(src_ws.protection)

    tgt_ws.freeze_panes = src_ws.freeze_panes

    # Spalten-/Zeilen-Dimensionen
    for col_key, dim in src_ws.column_dimensions.items():
        tgt_ws.column_dimensions[col_key].width = dim.width
        tgt_ws.column_dimensions[col_key].hidden = dim.hidden
        tgt_ws.column_dimensions[col_key].outline_level = dim.outline_level
        tgt_ws.column_dimensions[col_key].collapsed = dim.collapsed

    for row_idx, dim in src_ws.row_dimensions.items():
        tgt_ws.row_dimensions[row_idx].height = dim.height
        tgt_ws.row_dimensions[row_idx].hidden = dim.hidden
        tgt_ws.row_dimensions[row_idx].outline_level = dim.outline_level
        tgt_ws.row_dimensions[row_idx].collapsed = dim.collapsed

    # Zellen inkl. Styles
    for row in src_ws.iter_rows():
        for cell in row:
            tgt_cell = tgt_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                tgt_cell._style = copy_style(cell._style)
            tgt_cell.number_format = cell.number_format
            tgt_cell.protection = copy_style(cell.protection)
            tgt_cell.alignment = copy_style(cell.alignment)

    # Merges
    for merged_range in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged_range))

    return tgt_ws


def build_collection_workbook(period: str, suffix: str):
    """
    suffix: "_g" oder "_INTERN"
    baut:
      INSO_Land_<period>_SAMMEL_g.xlsx
      INSO_Land_<period>_SAMMEL_INTERN.xlsx
    """
    # Dateien pro Tabelle suchen
    def find_one(table_no):
        pattern = os.path.join(OUTPUT_DIR, f"Tabelle-{table_no}-Land_*{period}*{suffix}.xlsx")
        hits = sorted(glob.glob(pattern))
        return hits[0] if hits else None

    f1 = find_one(1)
    f2 = find_one(2)
    f3 = find_one(3)
    f5 = find_one(5)

    missing = [t for t, f in [(1, f1), (2, f2), (3, f3), (5, f5)] if f is None]
    if missing:
        print(f"[SAMMEL] Periode {period} ({suffix}): fehlende Dateien für Tabellen {missing} – Sammelmappe wird übersprungen.")
        return

    # Zielmappe anlegen
    out_wb = openpyxl.Workbook()
    # Default-Sheet löschen
    default = out_wb.active
    out_wb.remove(default)

    # Reihenfolge 1/2/3/5
    for path in [f1, f2, f3]:
        wb = openpyxl.load_workbook(path)
        ws = wb[wb.sheetnames[0]]
        copy_sheet_to_workbook(ws, out_wb, ws.title)

    # Tabelle 5: alle Blätter 1:1 übernehmen
    wb5 = openpyxl.load_workbook(f5)
    for ws in wb5.worksheets:
        copy_sheet_to_workbook(ws, out_wb, ws.title)

    # Dateiname
    tag = "g" if suffix == "_g" else "INTERN"
    out_path = os.path.join(OUTPUT_DIR, f"INSO_Land_{period}_SAMMEL_{tag}.xlsx")
    out_wb.save(out_path)
    print(f"[SAMMEL] erstellt: {out_path}")


def build_all_collections():
    """
    Findet alle Perioden, die in Ausgabedateien vorkommen,
    und erstellt pro Periode je eine Sammelmappe für _g und _INTERN.
    """
    files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
    periods = set()

    for f in files:
        if f.endswith("_g.xlsx") or f.endswith("_INTERN.xlsx"):
            p = detect_period_from_filename(f)
            if p:
                periods.add(p)

    for p in sorted(periods):
        build_collection_workbook(p, "_g")
        build_collection_workbook(p, "_INTERN")


# ------------------------------------------------------------
# Main
# ------------------------------------------------------------

def main():
    print("Starte Tabellen-Formatter (1,2,3,5 – INTERN & EXTERN)...")
    print(f"Arbeitsverzeichnis: {os.getcwd()}\n")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for table_no in (1, 2, 3, 5):
        pattern = f"Tabelle-{table_no}-Land_*.xlsx"
        candidates = sorted(glob.glob(pattern))

        raw_files = [
            f for f in candidates
            if not f.endswith("_g.xlsx") and not f.endswith("_INTERN.xlsx")
        ]

        if not raw_files:
            print(f"Keine Rohdateien für Tabelle {table_no} gefunden ({pattern}) – übersprungen.\n")
            continue

        tmpl_info = TEMPLATES.get(table_no)
        if not tmpl_info:
            print(f"[WARNUNG] Keine Vorlagenkonfiguration für Tabelle {table_no}.\n")
            continue

        tmpl_ext = os.path.join(LAYOUT_DIR, tmpl_info["ext"])
        tmpl_int = os.path.join(LAYOUT_DIR, tmpl_info["int"])

        for raw_path in raw_files:
            base_name = os.path.splitext(os.path.basename(raw_path))[0]
            is_jj = "-JJ" in base_name

            if table_no == 1:
                process_table1(raw_path, tmpl_ext, tmpl_int, is_jj)
            elif table_no in (2, 3):
                process_table2_or_3(table_no, raw_path, tmpl_ext, tmpl_int, is_jj)
            elif table_no == 5:
                process_table5(raw_path, tmpl_ext, tmpl_int, is_jj)

            print()

    # Sammelmappen pro Periode erzeugen
    print("Erzeuge Sammelmappen pro Periode (_g / _INTERN)...")
    build_all_collections()

    print("\nFertig. 🎉")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\nFEHLER AUFGETRETEN:")
        print(e)
    finally:
        print("\n--- Ende der Verarbeitung ---")
        input("Bitte Eingabetaste drücken, um das Fenster zu schließen...")


