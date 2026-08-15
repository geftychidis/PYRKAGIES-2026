#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Αναδημιουργεί το ΠΡΟΟΔΟΣ_ΣΥΛΛΟΓΗΣ_2026.xlsx από το LINKS_ΔΑΣΙΚΕΣ_ΠΥΡΚΑΓΙΕΣ_2026.md.

Παράγει μία γραμμή ανά ημερολογιακή ημέρα (1/5/2026 → σήμερα ή τελευταία ημέρα με
δεδομένα), με τον αριθμό συνδέσμων που φέρουν εκείνη την ημερομηνία, το σωρευτικό
σύνολο, και δύο γραφήματα (ημερήσια κατανομή + σωρευτική καμπύλη).

Χρήση:  python3 scripts/build_progress_xlsx.py [YYYY-MM-DD]
        (προαιρετικό όρισμα: τελευταία ημέρα του πίνακα· default = σήμερα)
"""

import collections
import datetime as dt
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
LINKS = ROOT / "LINKS_ΔΑΣΙΚΕΣ_ΠΥΡΚΑΓΙΕΣ_2026.md"
OUT = ROOT / "ΠΡΟΟΔΟΣ_ΣΥΛΛΟΓΗΣ_2026.xlsx"
SEASON_START = dt.date(2026, 5, 1)

MONTHS = ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος",
          "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος"]

HDR_FILL = PatternFill("solid", fgColor="1F4E5F")
BASE_FILL = PatternFill("solid", fgColor="EAEFF2")
THIN = Side(style="thin", color="B7C3C9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


YUTORI_MARK = "via Yutori Scout"


def parse_links():
    """Διαβάζει το LINKS και επιστρέφει τα πλήθη συνδέσμων ανά ημέρα και ανά πηγή.

    Returns:
        per_day:  {date: [λοιπές πηγές, via Yutori Scout]}
        undated:  [λοιπές πηγές, via Yutori Scout] για τις μόνιμες/ζωντανές πηγές
    """
    per_day = collections.defaultdict(lambda: [0, 0])
    undated = [0, 0]
    for line in LINKS.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|") or "http" not in line:
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 6:
            continue
        first = cells[0]
        if first in ("Ημ/νία", "Ημερομηνία ενημέρωσης") or set(first) <= set("-: "):
            continue
        idx = 1 if YUTORI_MARK in line else 0
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", first)
        if m:
            per_day[dt.date(*map(int, m.groups()))][idx] += 1
        else:
            undated[idx] += 1
    return per_day, undated


def build(per_day, undated, last_day):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ημερήσια Πρόοδος"

    headers = ["Ημερομηνία (μέρα/μήνας)", "Νέα Links ημέρας", "Λοιπές πηγές (web)",
               "via Yutori Scout", "Σύνολο Links"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    # Γραμμή βάσης: μόνιμες / ζωντανές πηγές χωρίς ημερομηνία περιστατικού.
    running = sum(undated)
    ws.append(["— (μόνιμες)", sum(undated), undated[0], undated[1], running])
    for c in ws[2]:
        c.fill = BASE_FILL
        c.border = BORDER

    first_data_row = 3
    day = SEASON_START
    while day <= last_day:
        other, yut = per_day.get(day, (0, 0))
        running += other + yut
        ws.append([day.strftime("%d/%m"), other + yut, other, yut, running])
        day += dt.timedelta(days=1)
    last_data_row = ws.max_row

    for row in ws.iter_rows(min_row=first_data_row, max_row=last_data_row, max_col=5):
        for c in row:
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")

    for col, w in zip("ABCDE", [22, 18, 20, 18, 14]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    # --- Γράφημα 1: ημερήσια κατανομή ανά πηγή (web vs Yutori Scout) ---
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.title = "Νέα links ανά ημέρα ανά πηγή — Δασικές Πυρκαγιές Ελλάδα 2026"
    bar.y_axis.title = "Αριθμός links"
    bar.x_axis.title = "Ημερομηνία"
    bar.height, bar.width = 9.5, 34
    bar.gapWidth = 40
    data = Reference(ws, min_col=3, max_col=4, min_row=1, max_row=last_data_row)
    cats = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.x_axis.tickLblSkip = 3
    bar.x_axis.tickMarkSkip = 3
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    ws.add_chart(bar, "G2")

    # --- Γράφημα 2: σωρευτική πορεία συλλογής ---
    line = LineChart()
    line.title = "Σωρευτικό σύνολο συνδέσμων"
    line.y_axis.title = "Σύνολο links"
    line.x_axis.title = "Ημερομηνία"
    line.height, line.width = 9.5, 34
    line.legend = None
    cum = Reference(ws, min_col=5, min_row=1, max_row=last_data_row)
    line.add_data(cum, titles_from_data=True)
    line.set_categories(cats)
    line.x_axis.tickLblSkip = 3
    line.x_axis.tickMarkSkip = 3
    line.x_axis.delete = False
    line.y_axis.delete = False
    ws.add_chart(line, "G22")

    # --- Φύλλο σύνοψης ανά μήνα ---
    ws2 = wb.create_sheet("Σύνοψη ανά μήνα")
    ws2.append(["Μήνας", "Links", "εκ των οποίων via Yutori", "Ημέρες με links",
                "Μέγιστο ημέρας", "Ημερομηνία μεγίστου"])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER
    by_month = collections.defaultdict(list)
    for d, (other, yut) in per_day.items():
        if d <= last_day:
            by_month[(d.year, d.month)].append((d, other + yut, yut))
    tot_links = tot_yut = 0
    for (y, m) in sorted(by_month):
        items = by_month[(y, m)]
        peak_day, peak_n, _ = max(items, key=lambda t: (t[1], t[0]))
        links = sum(n for _, n, _ in items)
        yut = sum(v for _, _, v in items)
        tot_links += links
        tot_yut += yut
        ws2.append([f"{MONTHS[m - 1]} {y}", links, yut, len(items),
                    peak_n, peak_day.strftime("%d/%m")])
    ws2.append(["Μόνιμες / ζωντανές πηγές", sum(undated), undated[1], "—", "—", "—"])
    ws2.append(["ΣΥΝΟΛΟ", tot_links + sum(undated), tot_yut + undated[1], "—", "—", "—"])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=6):
        for c in row:
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")
    for col in ("A", "B", "C"):
        ws2[f"{col}{ws2.max_row}"].font = Font(bold=True)
    for i, w in enumerate([26, 12, 22, 16, 16, 20], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(OUT)
    return last_data_row - first_data_row + 1, running


def main():
    last_day = (dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1
                else dt.date.today())
    per_day, undated = parse_links()
    if per_day:
        last_day = max(last_day, max(per_day))
    days, total = build(per_day, undated, last_day)
    yut = sum(v[1] for v in per_day.values()) + undated[1]
    print(f"OK: {days} ημέρες ({SEASON_START:%d/%m} → {last_day:%d/%m}), "
          f"σύνολο {total} links ({total - sum(undated)} με ημερομηνία + "
          f"{sum(undated)} μόνιμες) — εκ των οποίων {yut} via Yutori Scout")


if __name__ == "__main__":
    main()
