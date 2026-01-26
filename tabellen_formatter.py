import os
import glob
from copy import copy as copy_style
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment


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
    1: ("Tabelle-1-Layout_g.xlsx", "Tabelle-1-Layout_INTERN.xlsx"),
    2: ("Tabelle-2-Layout_g.xlsx", "Tabelle-2-Layout_INTERN.xlsx"),
    3: ("Tabelle-3-Layout_g.xlsx", "Tabelle-3-Layout_INTERN.xlsx"),
    5: ("Tabelle-5-Layout_g.xlsx", "Tabelle-5-Layout_INTERN.xlsx"),
}

INTERNAL_HEADER_TEXT = "NUR FÜR DEN INTERNEN DIENSTGEBRAUCH"


# ------------------------------------------------------------
# Helfer
# ------------------------------------------------------------

def is_numeric_like(v):
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "-", "X"):
            return False
        try:
            float(s.replace(".", "").replace(",", "."))
            return True
        except Exception:
            return False
    return False


def extract_stand_from_raw(ws):
    """
    Stand: dd.mm.yyyy steht i.d.R. in der letzten Zeile (Copyright-Zeile)
    Wir extrahieren den Text hinter 'Stand:'.
    """
    max_row = ws.max_row
    for r in range(max_row, max_row - 10, -1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "Stand:" in v:
                # z.B. "... Stand:15.12.2025"
                idx = v.find("Stand:")
                return v[idx:].strip()
    return ""


def extract_month_from_raw(ws_raw, table_no):
    """
    Monat/Periode in den Tabellen steht je nach Tabelle in einer Zeile.
    Wir nehmen hier simpel: erste Zeile, die wie 'Dezember 2025' aussieht.
    """
    for r in range(1, 20):
        v = ws_raw.cell(row=r, column=1).value
        if isinstance(v, str):
            s = v.strip()
            # sehr grob: enthält ein Jahr und ist nicht der lange Titel
            if any(m in s for m in ["Januar", "Februar", "März", "April", "Mai", "Juni",
                                    "Juli", "August", "September", "Oktober", "November", "Dezember"]) and "20" in s:
                return s
    return ""


def current_year():
    return datetime.now().year


def update_footer_with_stand_and_copyright(ws, stand_text):
    """
    In der letzten Zeile soll stehen:
    (C)opyright YYYY Bayerisches Landesamt für Statistik    Stand:xx.xx.xxxx
    Wenn keine Copyrightzeile existiert: am Ende 1 Leerzeile + Copyrightzeile einfügen.
    """
    yr = current_year()
    target_row = ws.max_row

    # Suche nach bestehender Copyright-Zeile
    found_row = None
    for r in range(ws.max_row, max(ws.max_row - 15, 1), -1):
        row_texts = []
        for c in range(1, min(ws.max_column, 20) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                row_texts.append(v)
        joined = " ".join(row_texts)
        if "Bayerisches Landesamt für Statistik" in joined or "(C)opyright" in joined or "Copyright" in joined:
            found_row = r
            break

    if found_row is None:
        # 1 Leerzeile + neue Zeile
        ws.append([])
        ws.append([])
        found_row = ws.max_row

    # Wir schreiben die Copyright-Zeile in Spalte A
    ws.cell(row=found_row, column=1).value = f"(C)opyright {yr} Bayerisches Landesamt für Statistik"
    # Stand immer in die letzte Spalte, rechtsbündig, gleiche Schrift wie A
    last_col = ws.max_column
    stand_cell = ws.cell(row=found_row, column=last_col)
    stand_cell.value = stand_text.replace("Stand:", "Stand:").strip()

    # Format Stand-Zelle
    stand_cell.alignment = Alignment(horizontal="right", vertical="bottom")

    # Falls irgendwo doppelt "Stand:" drinsteht: bereinigen
    if isinstance(stand_cell.value, str):
        s = stand_cell.value
        # wenn "Stand:" doppelt vorkommt, reduziere
        if s.count("Stand:") > 1:
            first = s.find("Stand:")
            stand_cell.value = "Stand:" + s[first + len("Stand:"):].strip()


def get_merged_secondary_checker(ws):
    """
    Liefert Funktion is_secondary_cell(r,c), die True zurückgibt,
    wenn (r,c) innerhalb eines Merge-Range ist, aber NICHT die Top-Left-Zelle.
    """
    merged = list(ws.merged_cells.ranges)

    def is_secondary_cell(r, c):
        for rng in merged:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                return not (r == rng.min_row and c == rng.min_col)
        return False

    return is_secondary_cell


# ------------------------------------------------------------
# Tabelle 1 (bereits stabil bei dir – hier unverändert gelassen)
# ------------------------------------------------------------

def process_table1(raw_path, tmpl_ext_path, tmpl_int_path):
    print(f"Verarbeite Tabelle 1 aus '{raw_path}' ...")

    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    month_text = extract_month_from_raw(ws_raw, 1)
    stand_text = extract_stand_from_raw(ws_raw)

    # EXTERN
    if os.path.exists(tmpl_ext_path):
        wb_ext = openpyxl.load_workbook(tmpl_ext_path)
        ws_ext = wb_ext[wb_ext.sheetnames[0]]

        # Monat extern: Zeile 3
        ws_ext.cell(row=3, column=1).value = month_text

        # Datenbereich grob: alles numerische ab Spalte 2
        is_sec = get_merged_secondary_checker(ws_ext)
        for r in range(1, ws_ext.max_row + 1):
            for c in range(2, ws_ext.max_column + 1):
                if is_sec(r, c):
                    continue
                ws_ext.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

        update_footer_with_stand_and_copyright(ws_ext, stand_text)

        out_ext = raw_path.replace(".xlsx", "_g.xlsx")
        wb_ext.save(out_ext)
        print(f"  -> Extern: {out_ext}")
    else:
        print(f"  [WARNUNG] Vorlage extern für Tabelle 1 nicht gefunden: {tmpl_ext_path}")

    # INTERN
    if os.path.exists(tmpl_int_path):
        wb_int = openpyxl.load_workbook(tmpl_int_path)
        ws_int = wb_int[wb_int.sheetnames[0]]

        # Kopfzeile
        ws_int.cell(row=1, column=1).value = INTERNAL_HEADER_TEXT
        # Monat intern: Zeile 6
        ws_int.cell(row=6, column=1).value = month_text

        is_sec = get_merged_secondary_checker(ws_int)
        for r in range(1, ws_int.max_row + 1):
            for c in range(2, ws_int.max_column + 1):
                if is_sec(r, c):
                    continue
                ws_int.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

        update_footer_with_stand_and_copyright(ws_int, stand_text)

        out_int = raw_path.replace(".xlsx", "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle 1 nicht gefunden: {tmpl_int_path}")


# ------------------------------------------------------------
# Verarbeitung für Tabelle 2 & 3 (FIX: Startspalte = 2 statt 3!)
# ------------------------------------------------------------

def process_table2_or_3(table_no, raw_path, tmpl_ext_path, tmpl_int_path):
    print(f"Verarbeite Tabelle {table_no} aus '{raw_path}' ...")

    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]]

    month_text = extract_month_from_raw(ws_raw, table_no)
    stand_text = extract_stand_from_raw(ws_raw)

    # FIX: Bei Tabelle 2 & 3 stehen Werte schon in Spalte B -> start_col = 2
    start_col = 2

    def fill_numeric(ws_t, ws_raw):
        is_sec = get_merged_secondary_checker(ws_t)
        max_row_t = ws_t.max_row
        max_col_t = ws_t.max_column

        # Daten- & Fußnotenbereich finden
        def detect_data_and_footer(ws, numeric_col=start_col):
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

        fdr_raw, ft_raw = detect_data_and_footer(ws_raw, numeric_col=start_col)
        fdr_t, ft_t = detect_data_and_footer(ws_t, numeric_col=start_col)

        n_rows_raw = ft_raw - fdr_raw
        n_rows_t = ft_t - fdr_t
        n_rows = min(n_rows_raw, n_rows_t)

        for offset in range(n_rows):
            r_raw = fdr_raw + offset
            r_t = fdr_t + offset

            # FIX: ab start_col (2) überschreiben
            for c in range(start_col, max_col_t + 1):
                if is_sec(r_t, c):
                    continue
                ws_t.cell(row=r_t, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    # ----- EXTERN -----
    if os.path.exists(tmpl_ext_path):
        wb_ext = openpyxl.load_workbook(tmpl_ext_path)
        ws_ext = wb_ext[wb_ext.sheetnames[0]]

        # Monat extern: Zeile 3, Spalte 1
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

        # Kopfzeile intern
        ws_int.cell(row=1, column=1).value = INTERNAL_HEADER_TEXT

        # Monat intern: Zeile 6, Spalte 1
        ws_int.cell(row=6, column=1).value = month_text

        fill_numeric(ws_int, ws_raw)
        update_footer_with_stand_and_copyright(ws_int, stand_text)

        out_int = raw_path.replace(".xlsx", "_INTERN.xlsx")
        wb_int.save(out_int)
        print(f"  -> Intern: {out_int}")
    else:
        print(f"  [WARNUNG] Vorlage intern für Tabelle {table_no} nicht gefunden: {tmpl_int_path}")


# ------------------------------------------------------------
# Tabelle 5 (hier nur Platzhalter – bei dir war das bereits separat stabil)
# Wenn du willst, kann ich deinen aktuellen Table5-Teil hier wieder einbauen.
# ------------------------------------------------------------

def process_table5(raw_path, tmpl_ext_path, tmpl_int_path):
    # HINWEIS: Table5 ist bei dir schon "perfekt" gewesen – hier nicht erneut vollständig
    # eingebaut, um nichts zu zerstören. Wenn du mir sagst "baue Table5 wieder rein",
    # poste ich die vollständige integrierte Version.
    print("Tabelle 5: in dieser Datei aktuell nicht enthalten (bitte deinen stabilen Table5-Code einfügen).")


# ------------------------------------------------------------
# Main
# ------------------------------------------------------------

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Suche nach Eingabedateien Tabelle-1/2/3/5
    raw_files = sorted(glob.glob("Tabelle-*-Land_*.xlsx"))
    raw_files = [f for f in raw_files if not (f.endswith("_g.xlsx") or f.endswith("_INTERN.xlsx"))]

    if not raw_files:
        print("Keine Eingabedateien gefunden (Tabelle-*-Land_*.xlsx).")
        return

    for raw_path in raw_files:
        base = os.path.basename(raw_path)

        if base.startswith("Tabelle-1-"):
            tmpl_ext, tmpl_int = TEMPLATES[1]
            process_table1(
                raw_path,
                os.path.join(LAYOUT_DIR, tmpl_ext),
                os.path.join(LAYOUT_DIR, tmpl_int),
            )

        elif base.startswith("Tabelle-2-"):
            tmpl_ext, tmpl_int = TEMPLATES[2]
            process_table2_or_3(
                2,
                raw_path,
                os.path.join(LAYOUT_DIR, tmpl_ext),
                os.path.join(LAYOUT_DIR, tmpl_int),
            )

        elif base.startswith("Tabelle-3-"):
            tmpl_ext, tmpl_int = TEMPLATES[3]
            process_table2_or_3(
                3,
                raw_path,
                os.path.join(LAYOUT_DIR, tmpl_ext),
                os.path.join(LAYOUT_DIR, tmpl_int),
            )

        elif base.startswith("Tabelle-5-"):
            tmpl_ext, tmpl_int = TEMPLATES[5]
            process_table5(
                raw_path,
                os.path.join(LAYOUT_DIR, tmpl_ext),
                os.path.join(LAYOUT_DIR, tmpl_int),
            )

    print("Fertig.")


if __name__ == "__main__":
    main()
