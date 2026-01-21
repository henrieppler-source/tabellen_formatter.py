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

# Basisordner für Layouts und Ausgabedateien (relativ zum aktuellen Arbeitsverzeichnis)
LAYOUT_DIR = "Layouts"
OUTPUT_DIR = "Ausgabedateien"

# Zuordnung: Tabellennummer -> Rohblattname
RAW_SHEET_NAMES = {
    1: "XML-Tab1-Land",
    2: "XML-Tab2-Land",
    3: "XML-Tab3-Land",
    5: "XML-Tab5-Land",
}

# Zuordnung: Tabellennummer -> Layout-Dateien (nur Dateinamen, Pfad kommt über LAYOUT_DIR)
TEMPLATES = {
    1: {
        "ext": "Tabelle-1-Layout_g.xlsx",
        "int": "Tabelle-1-Layout_INTERN.xlsx",
    },
    2: {
        "ext": "Tabelle-2-Layout_g.xlsx",
        "int": "Tabelle-2-Layout_INTERN.xlsx",
    },
    3: {
        "ext": "Tabelle-3-Layout_g.xlsx",
        "int": "Tabelle-3-Layout_INTERN.xlsx",
    },
    5: {
        "ext": "Tabelle-5-Layout_g.xlsx",
        "int": "Tabelle-5-Layout_INTERN.xlsx",
    },
}


# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

def is_numeric_like(v):
    """Erkennt Zahlen/Platzhalter in Zellen (inkl. '-', 'X', mit Punkt/Komma)."""
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip().replace(".", "").replace(",", "")
        return s.isdigit() or v.strip() in ["-", "X"]
    return False


def extract_month_from_raw(ws, table_no):
    """
    Holt den Monats-/Zeitraum-Text aus dem Rohblatt.
    Position je nach Tabelle:
      - 1: A3
      - 2: A4
      - 3: A4
      - 5: A3
    """
    if table_no == 1:
        return ws.cell(row=3, column=1).value
    elif table_no in (2, 3):
        return ws.cell(row=4, column=1).value
    elif table_no == 5:
        return ws.cell(row=3, column=1).value
    else:
        return None


def extract_stand_from_raw(ws, max_search_rows=40):
    """Sucht den 'Stand:'-Text in den letzten Zeilen des Rohblatts."""
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(max_row, max(max_row - max_search_rows, 1) - 1, -1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "Stand:" in v:
                return v.strip()
    return None


def update_footer_with_stand_and_copyright(ws, stand_text):
    """
    Aktualisiert in einem Vorlagenblatt:
      - Copyright-Jahr auf aktuelles Jahr
      - Stand-Text in der ursprünglichen Stand-Spalte der Copyright-Zeile
    """
    max_row = ws.max_row
    max_col = ws.max_column
    current_year = datetime.now().year

    # Copyright-Zeile finden
    copyright_row = None
    for r in range(max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            text = v
            # Jahreszahl aktualisieren
            def repl(m):
                return f"(C)opyright {current_year}"
            new_text = re.sub(r"\(C\)opyright\s+\d{4}", repl, text)
            ws.cell(row=r, column=1).value = new_text
            copyright_row = r
            break

    if not copyright_row or not stand_text:
        return

    # Spalte ermitteln, in der ursprünglich 'Stand:' stand (falls vorhanden)
    stand_col = None
    for c in range(1, max_col + 1):
        v = ws.cell(row=copyright_row, column=c).value
        if isinstance(v, str) and "Stand:" in v:
            stand_col = c
            break

    # Alle anderen 'Stand:'-Vorkommen im Blatt löschen
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().startswith("Stand:") and r != copyright_row:
                ws.cell(row=r, column=c).value = ""

    if stand_col is None:
        stand_col = max_col  # Fallback: letzte Spalte

    # Stil von Copyright-Zelle übernehmen
    cop_cell = ws.cell(row=copyright_row, column=1)
    tgt = ws.cell(row=copyright_row, column=stand_col)
    tgt.value = stand_text

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
    """
    Liefert eine Funktion, die sagt, ob eine Zelle sekundärer Teil
    eines Merge-Bereichs ist (damit wir sie nicht überschreiben).
    """
    merged = list(ws.merged_cells.ranges)

    def is_secondary(row, col):
        for rg in merged:
            if rg.min_row <= row <= rg.max_row and rg.min_col <= col <= rg.max_col:
                return not (row == rg.min_row and col == rg.min_col)
        return False

    return is_secondary


def mark_cells_with_1_or_2(ws, col_index, fill):
    """Markiert Zellen in gegebener Spalte, falls Wert 1 oder 2 ist."""
    max_row = ws.max_row
    for r in range(1, max_row + 1):
        cell = ws.cell(row=r, column=col_index)
        v = cell.value
        if isinstance(v, (int, float)) and v in (1, 2):
            cell.fill = fill
        elif isinstance(v, str) and v.strip() in ("1", "2"):
            cell.fill = fill


def format_numeric_cells(ws):
    """
    Setzt für alle numerischen Zellen ein einheitliches Zahlenformat:
    Ganzzahl mit Leerzeichen als Tausendertrennzeichen, keine Dezimalstellen.
    0 bleibt 0, '-' (Text) bleibt unberührt.
    """
    thousands_format = "# ##0"
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, (int, float)):
                # Floats auf ganze Zahl runden
                if isinstance(v, float):
                    if v.is_integer():
                        cell.value = int(v)
                    else:
                        cell.value = int(round(v))
                cell.number_format = thousands_format


# ------------------------------------------------------------
# Verarbeitung für Tabelle 1
# ------------------------------------------------------------

def build_table1_workbook(raw_path, template_path, internal_layout):
    """Erzeugt eine Tabelle-1-Arbeitsmappe auf Basis einer Vorlage."""
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    month_text = extract_month_from_raw(ws_raw, 1)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]

    # Monat setzen
    if internal_layout:
        ws.cell(row=5, column=1).value = month_text  # intern
    else:
        ws.cell(row=3, column=1).value = month_text  # extern

    is_sec = get_merged_secondary_checker(ws)
    max_row_ws = ws.max_row
    max_col_ws = ws.max_column

    # Daten- und Fußnotenbereich in Roh- und Vorlagenblatt bestimmen
    def detect_data_and_footer(sheet, numeric_col=4):
        max_row = sheet.max_row
        first_data = None
        for r in range(1, max_row + 1):
            v = sheet.cell(row=r, column=numeric_col).value
            if is_numeric_like(v):
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

    n_rows_raw = ft_raw - fdr_raw
    n_rows_t = ft_t - fdr_t
    n_rows = min(n_rows_raw, n_rows_t)

    for offset in range(n_rows):
        r_raw = fdr_raw + offset
        r_t = fdr_t + offset
        for c in range(1, max_col_ws + 1):
            if is_sec(r_t, c):
                continue
            ws.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws, stand_text)
    format_numeric_cells(ws)
    return wb


def process_table1(raw_path, tmpl_ext_path, tmpl_int_path, is_jj):
    print(f"Verarbeite Tabelle 1 aus '{raw_path}' ...")

    # INTERN immer auf Basis der internen Vorlage
    wb_int = None
    if os.path.exists(tmpl_int_path):
        wb_int = build_table1_workbook(raw_path, tmpl_int_path, internal_layout=True)
        base = os.path.splitext(os.path.basename(raw_path))[0]
        out_int = os.path.join(OUTPUT_DIR, base + "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle 1 nicht gefunden: {tmpl_int_path}")

    # EXTERN
    if is_jj:
        # JJ: extern identisch zu intern, aber ohne erste Zeile + Markierungen
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table1_workbook(raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            ws_ext = wb_ext[wb_ext.sheetnames[0]]
            # Erste Zeile (Hinweis 'Nur für...') entfernen
            ws_ext.cell(row=1, column=1).value = None

            # Markierungen in Spalte G (7)
            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            mark_cells_with_1_or_2(ws_ext, 7, fill)

            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern (JJ): {out_ext}")
    else:
        # normaler Monat: externe Layoutvorlage
        if os.path.exists(tmpl_ext_path):
            wb_ext = build_table1_workbook(raw_path, tmpl_ext_path, internal_layout=False)
            base = os.path.splitext(os.path.basename(raw_path))[0]
            out_ext = os.path.join(OUTPUT_DIR, base + "_g.xlsx")
            wb_ext.save(out_ext)
            print(f"  -> Extern: {out_ext}")
        else:
            print(f"  [WARNUNG] Vorlage extern für Tabelle 1 nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Verarbeitung für Tabelle 2 & 3 (ähnliche Struktur)
# ------------------------------------------------------------

def build_table2_or_3_workbook(table_no, raw_path, template_path, internal_layout):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]]

    month_text = extract_month_from_raw(ws_raw, table_no)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]

    # Monat setzen
    if internal_layout:
        # intern: Zeile 6, Spalte 1
        ws.cell(row=6, column=1).value = month_text
    else:
        # extern: Zeile 3, Spalte 1
        ws.cell(row=3, column=1).value = month_text

    is_sec = get_merged_secondary_checker(ws)
    max_row_t = ws.max_row
    max_col_t = ws.max_column

    # Daten- & Fußnotenbereich
    def detect_data_and_footer(sheet, numeric_col=3):
        max_row = sheet.max_row
        first_data = None
        for r in range(1, max_row + 1):
            v = sheet.cell(row=r, column=numeric_col).value
            if is_numeric_like(v):
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

    n_rows_raw = ft_raw - fdr_raw
    n_rows_t = ft_t - fdr_t
    n_rows = min(n_rows_raw, n_rows_t)

    for offset in range(n_rows):
        r_raw = fdr_raw + offset
        r_t = fdr_t + offset
        # Nur numerische Spalten (>=3) überschreiben, Text/Fußnoten in A/B bleiben
        for c in range(3, max_col_t + 1):
            if is_sec(r_t, c):
                continue
            ws.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws, stand_text)
    format_numeric_cells(ws)
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
        print(f"  [WARNUNG] Vorlage intern für Tabelle {table_no} nicht gefunden: {tmpl_int_path}")

    if is_jj:
        # JJ: extern identisch intern, ohne erste Zeile + Markierungen
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table2_or_3_workbook(table_no, raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            ws_ext = wb_ext[wb_ext.sheetnames[0]]
            ws_ext.cell(row=1, column=1).value = None

            # Markierungen: Tabelle 2 & 3 -> Spalte E (5)
            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            mark_cells_with_1_or_2(ws_ext, 5, fill)

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
            print(f"  [WARNUNG] Vorlage extern für Tabelle {table_no} nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Verarbeitung für Tabelle 5 (5 Blöcke / 5 Blätter)
# ------------------------------------------------------------

def build_table5_workbook(raw_path, template_path, internal_layout):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[5]]

    month_text = extract_month_from_raw(ws_raw, 5)
    stand_text = extract_stand_from_raw(ws_raw)

    wb = openpyxl.load_workbook(template_path)
    max_row = ws_raw.max_row

    # Blöcke finden via "Bayern x)" in Spalte B
    starts = []
    for r in range(1, max_row + 1):
        v = ws_raw.cell(row=r, column=2).value
        if isinstance(v, str) and re.match(r"Bayern\s+\d\)", v.strip()):
            starts.append(r)

    block_ranges = []
    for i, start in enumerate(starts):
        if i < len(starts) - 1:
            end = starts[i + 1] - 1
        else:
            end = max_row
        # Trailing Leerzeilen im Block wegschneiden
        last_nonempty = start
        for r in range(start, end + 1):
            if any(ws_raw.cell(row=r, column=c).value not in (None, "") for c in range(1, 10)):
                last_nonempty = r
        block_ranges.append((start, last_nonempty))

    # Gemeinsame Funktion: Füllt eine Vorlage-Tabelle aus einem Block
    def fill_sheet_from_block(ws_t, start_row, end_row):
        is_sec = get_merged_secondary_checker(ws_t)
        max_row_t = ws_t.max_row

        # erste Datenzeile im Template: erste Zeile mit numerischem Wert in Spalte 3
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
            for c in range(3, 9):  # nur C..H überschreiben (Zahlen)
                if is_sec(t_r, c):
                    continue
                ws_t.cell(row=t_r, column=c).value = ws_raw.cell(row=raw_r, column=c).value
            raw_r += 1
            t_r += 1

    for i, (start, end) in enumerate(block_ranges):
        if i >= len(wb.worksheets):
            break
        ws_t = wb.worksheets[i]

        # Monat setzen
        if internal_layout:
            # intern: Zeile 5, Spalte 1
            ws_t.cell(row=5, column=1).value = month_text
        else:
            # extern: Zeile 3, Spalte 1
            ws_t.cell(row=3, column=1).value = month_text

        fill_sheet_from_block(ws_t, start, end)
        update_footer_with_stand_and_copyright(ws_t, stand_text)
        format_numeric_cells(ws_t)

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
        print(f"  [WARNUNG] Vorlage intern für Tabelle 5 nicht gefunden: {tmpl_int_path}")

    if is_jj:
        # JJ: extern identisch intern, ohne erste Zeile + Markierungen in Spalte F (6) auf allen Blättern
        if wb_int is None and os.path.exists(tmpl_int_path):
            wb_int = build_table5_workbook(raw_path, tmpl_int_path, internal_layout=True)
        if wb_int is not None:
            wb_ext = wb_int
            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            for ws_ext in wb_ext.worksheets:
                # Erste Zeile mit 'Nur für...' entfernen
                ws_ext.cell(row=1, column=1).value = None
                # Markierungen in Spalte F (6)
                mark_cells_with_1_or_2(ws_ext, 6, fill)

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
            print(f"  [WARNUNG] Vorlage extern für Tabelle 5 nicht gefunden: {tmpl_ext_path}")


# ------------------------------------------------------------
# Hauptprogramm
# ------------------------------------------------------------

def main():
    print("Starte Tabellen-Formatter (1,2,3,5 – INTERN & EXTERN)...")
    cwd = os.getcwd()
    print(f"Arbeitsverzeichnis: {cwd}")
    print()

    # Ausgabeverzeichnis sicherstellen
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    table_numbers = [1, 2, 3, 5]

    for table_no in table_numbers:
        pattern = f"Tabelle-{table_no}-Land_*.xlsx"
        candidates = sorted(glob.glob(pattern))

        # Nur Rohdateien, keine bereits erzeugten _g / _INTERN
        raw_files = [
            f for f in candidates
            if not f.endswith("_g.xlsx") and not f.endswith("_INTERN.xlsx")
        ]

        if not raw_files:
            print(f"Keine Rohdateien für Tabelle {table_no} gefunden ({pattern}) – wird übersprungen.")
            print()
            continue

        for raw_path in raw_files:
            tmpl_info = TEMPLATES.get(table_no)
            if not tmpl_info:
                print(f"[WARNUNG] Keine Vorlagenkonfiguration für Tabelle {table_no} vorhanden.")
                continue

            tmpl_ext = os.path.join(LAYOUT_DIR, tmpl_info["ext"])
            tmpl_int = os.path.join(LAYOUT_DIR, tmpl_info["int"])

            # JJ-Sonderfall: Dateiname enthält '-JJ' vor der Endung
            base_name = os.path.splitext(os.path.basename(raw_path))[0]
            is_jj = "-JJ" in base_name

            if table_no == 1:
                process_table1(raw_path, tmpl_ext, tmpl_int, is_jj)
            elif table_no in (2, 3):
                process_table2_or_3(table_no, raw_path, tmpl_ext, tmpl_int, is_jj)
            elif table_no == 5:
                process_table5(raw_path, tmpl_ext, tmpl_int, is_jj)

            print()

    print("Fertig. 🎉")


if __name__ == "__main__":
    main()
