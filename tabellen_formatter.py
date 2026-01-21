import os
import glob
from copy import copy as copy_style
from datetime import datetime
import re

import openpyxl
from openpyxl.styles import Alignment


# ------------------------------------------------------------
# Konfiguration
# ------------------------------------------------------------

# Zuordnung: Tabellennummer -> Rohblattname
RAW_SHEET_NAMES = {
    1: "XML-Tab1-Land",
    2: "XML-Tab2-Land",
    3: "XML-Tab3-Land",
    5: "XML-Tab5-Land",
}

# Zuordnung: Tabellennummer -> Vorlagen-Dateien (extern / intern)
# -> Diese Dateien dienen als Layout-Vorlagen und sollten dauerhaft im Ordner liegen.
TEMPLATES = {
    1: {
        "ext": "Tabelle-1-Land_2025-11_g.xlsx",
        "int": "Tabelle-1-Land_2025-11_INTERN.xlsx",
    },
    2: {
        "ext": "Tabelle-2-Land_2025-11_g.xlsx",
        "int": "Tabelle-2-Land_2025-11_INTERN.xlsx",
    },
    3: {
        "ext": "Tabelle-3-Land_2025-11_g.xlsx",
        "int": "Tabelle-3-Land_2025-11_INTERN.xlsx",
    },
    5: {
        "ext": "Tabelle-5-Land_2025-11_g.xlsx",
        "int": "Tabelle-5-Land_2025-11_INTERN.xlsx",
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
            import re
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


# ------------------------------------------------------------
# Verarbeitung für Tabelle 1
# ------------------------------------------------------------

def process_table1(raw_path, tmpl_ext_path, tmpl_int_path):
    print(f"Verarbeite Tabelle 1 aus '{raw_path}' ...")

    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    month_text = extract_month_from_raw(ws_raw, 1)
    stand_text = extract_stand_from_raw(ws_raw)

    # ----- EXTERN -----
    if os.path.exists(tmpl_ext_path):
        wb_ext = openpyxl.load_workbook(tmpl_ext_path)
        ws_ext = wb_ext[wb_ext.sheetnames[0]]

        # Monat (extern): Zeile 3, Spalte 1
        ws_ext.cell(row=3, column=1).value = month_text

        is_sec = get_merged_secondary_checker(ws_ext)
        max_row_ext = ws_ext.max_row
        max_col_ext = ws_ext.max_column

        # Daten- und Fußnotenbereich in Roh- und Vorlagenblatt bestimmen
        def detect_data_and_footer(ws, numeric_col=4):
            max_row = ws.max_row
            first_data = None
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=numeric_col).value
                if is_numeric_like(v):
                    first_data = r
                    break
            if first_data is None:
                first_data = 1
            footnote_start = max_row + 1
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=1).value
                if isinstance(v, str) and v.strip().startswith("-"):
                    footnote_start = r
                    break
            return first_data, footnote_start

        fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=4)
        fdr_ext, ft_ext = detect_data_and_footer(ws_ext, numeric_col=4)

        n_rows_raw = ft_raw - fdr_raw
        n_rows_ext = ft_ext - fdr_ext
        n_rows = min(n_rows_raw, n_rows_ext)

        for offset in range(n_rows):
            r_raw = fdr_raw + offset
            r_ext = fdr_ext + offset
            for c in range(1, max_col_ext + 1):
                if is_sec(r_ext, c):
                    continue
                ws_ext.cell(row=r_ext, column=c).value = ws_raw.cell(row=r_raw, column=c).value

        update_footer_with_stand_and_copyright(ws_ext, stand_text)

        out_ext = raw_path.replace(".xlsx", "_g.xlsx")
        wb_ext.save(out_ext)
        print(f"  -> Extern: {out_ext}")
    else:
        print(f"  [WARNUNG] Vorlage extern für Tabelle 1 nicht gefunden: {tmpl_ext_path}")

    # ----- INTERN -----
    if os.path.exists(tmpl_int_path):
        wb_int = openpyxl.load_workbook(tmpl_int_path)
        ws_int = wb_int[wb_int.sheetnames[0]]

        # Monat (intern): Zeile 5, Spalte 1
        ws_int.cell(row=5, column=1).value = month_text

        is_sec_int = get_merged_secondary_checker(ws_int)
        max_row_int = ws_int.max_row
        max_col_int = ws_int.max_column

        def detect_data_and_footer(ws, numeric_col=4):
            max_row = ws.max_row
            first_data = None
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=numeric_col).value
                if is_numeric_like(v):
                    first_data = r
                    break
            if first_data is None:
                first_data = 1
            footnote_start = max_row + 1
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=1).value
                if isinstance(v, str) and v.strip().startswith("-"):
                    footnote_start = r
                    break
            return first_data, footnote_start

        fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=4)
        fdr_int, ft_int = detect_data_and_footer(ws_int, numeric_col=4)

        n_rows_raw = ft_raw - fdr_raw
        n_rows_int = ft_int - fdr_int
        n_rows = min(n_rows_raw, n_rows_int)

        for offset in range(n_rows):
            r_raw = fdr_raw + offset
            r_int = fdr_int + offset
            for c in range(1, max_col_int + 1):
                if is_sec_int(r_int, c):
                    continue
                ws_int.cell(row=r_int, column=c).value = ws_raw.cell(row=r_raw, column=c).value

        update_footer_with_stand_and_copyright(ws_int, stand_text)

        out_int = raw_path.replace(".xlsx", "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle 1 nicht gefunden: {tmpl_int_path}")


# ------------------------------------------------------------
# Verarbeitung für Tabelle 2 & 3 (ähnliche Struktur)
# ------------------------------------------------------------

def process_table2_or_3(table_no, raw_path, tmpl_ext_path, tmpl_int_path):
    print(f"Verarbeite Tabelle {table_no} aus '{raw_path}' ...")

    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]]

    month_text = extract_month_from_raw(ws_raw, table_no)
    stand_text = extract_stand_from_raw(ws_raw)

    # Gemeinsame Funktion für Datenkopie (nur numerische Spalten ab Spalte 3)
    def fill_numeric(ws_t, ws_raw):
        is_sec = get_merged_secondary_checker(ws_t)
        max_row_t = ws_t.max_row
        max_col_t = ws_t.max_column

        # Daten- & Fußnotenbereich
        def detect_data_and_footer(ws, numeric_col=3):
            max_row = ws.max_row
            first_data = None
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=numeric_col).value
                if is_numeric_like(v):
                    first_data = r
                    break
            if first_data is None:
                first_data = 1
            footnote_start = max_row + 1
            for r in range(1, max_row + 1):
                v = ws.cell(row=r, column=1).value
                if isinstance(v, str) and v.strip().startswith("-"):
                    footnote_start = r
                    break
            return first_data, footnote_start

        fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=3)
        fdr_t, ft_t = detect_data_and_footer(ws_t, numeric_col=3)

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
                ws_t.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    # ----- EXTERN -----
    if os.path.exists(tmpl_ext_path):
        wb_ext = openpyxl.load_workbook(tmpl_ext_path)
        ws_ext = wb_ext[wb_ext.sheetnames[0]]

        # Monat (extern): Zeile 3, Spalte 1
        ws_ext.cell(row=3, column=1).value = month_text

        fill_numeric(ws_ext, ws_raw)
        update_footer_with_stand_and_copyright(ws_ext, stand_text)

        out_ext = raw_path.replace(".xlsx", "_g.xlsx")
        wb_ext.save(out_ext)
        print(f"  -> Extern: {out_ext}")
    else:
        print(f"  [WARNUNG] Vorlage extern für Tabelle {table_no} nicht gefunden: {tmpl_ext_path}")

    # ----- INTERN -----
    if os.path.exists(tmpl_int_path):
        wb_int = openpyxl.load_workbook(tmpl_int_path)
        ws_int = wb_int[wb_int.sheetnames[0]]

        # Monat (intern): Zeile 6, Spalte 1
        ws_int.cell(row=6, column=1).value = month_text

        fill_numeric(ws_int, ws_raw)
        update_footer_with_stand_and_copyright(ws_int, stand_text)

        out_int = raw_path.replace(".xlsx", "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle {table_no} nicht gefunden: {tmpl_int_path}")


# ------------------------------------------------------------
# Verarbeitung für Tabelle 5 (5 Blöcke / 5 Blätter)
# ------------------------------------------------------------

def process_table5(raw_path, tmpl_ext_path, tmpl_int_path):
    print(f"Verarbeite Tabelle 5 aus '{raw_path}' ...")

    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[5]]

    month_text = extract_month_from_raw(ws_raw, 5)
    stand_text = extract_stand_from_raw(ws_raw)

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

    # ----- EXTERN (5.x-Blätter) -----
    if os.path.exists(tmpl_ext_path):
        wb_ext = openpyxl.load_workbook(tmpl_ext_path)

        for i, (start, end) in enumerate(block_ranges):
            if i >= len(wb_ext.worksheets):
                break
            ws_t = wb_ext.worksheets[i]

            # Monat extern: Zeile 3, Spalte 1
            ws_t.cell(row=3, column=1).value = month_text

            fill_sheet_from_block(ws_t, start, end)
            update_footer_with_stand_and_copyright(ws_t, stand_text)

        out_ext = raw_path.replace(".xlsx", "_g.xlsx")
        wb_ext.save(out_ext)
        print(f"  -> Extern: {out_ext}")
    else:
        print(f"  [WARNUNG] Vorlage extern für Tabelle 5 nicht gefunden: {tmpl_ext_path}")

    # ----- INTERN (5.x-Blätter) -----
    if os.path.exists(tmpl_int_path):
        wb_int = openpyxl.load_workbook(tmpl_int_path)

        for i, (start, end) in enumerate(block_ranges):
            if i >= len(wb_int.worksheets):
                break
            ws_t = wb_int.worksheets[i]

            # Monat intern: Zeile 5, Spalte 1
            ws_t.cell(row=5, column=1).value = month_text

            fill_sheet_from_block(ws_t, start, end)
            update_footer_with_stand_and_copyright(ws_t, stand_text)

        out_int = raw_path.replace(".xlsx", "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle 5 nicht gefunden: {tmpl_int_path}")


# ------------------------------------------------------------
# Hauptprogramm
# ------------------------------------------------------------

def main():
    print("Starte Tabellen-Formatter (1,2,3,5 – INTERN & EXTERN)...")
    cwd = os.getcwd()
    print(f"Arbeitsverzeichnis: {cwd}")
    print()

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

            tmpl_ext = tmpl_info["ext"]
            tmpl_int = tmpl_info["int"]

            if table_no == 1:
                process_table1(raw_path, tmpl_ext, tmpl_int)
            elif table_no in (2, 3):
                process_table2_or_3(table_no, raw_path, tmpl_ext, tmpl_int)
            elif table_no == 5:
                process_table5(raw_path, tmpl_ext, tmpl_int)

            print()

    print("Fertig. 🎉")


if __name__ == "__main__":
    main()

