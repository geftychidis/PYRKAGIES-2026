#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Αποκαθιστά (ή ξαναχτίζει) το φύλλο «Dashboard» στο βιβλίο ΠΥΡΚΑΓΙΕΣ 2026.

Γιατί υπάρχει: ο καθημερινός αυτοματισμός ξαναγράφει το .xlsx και, όταν το κάνει
με βιβλιοθήκη που δεν καταλαβαίνει γραφήματα (π.χ. openpyxl/pandas), το φύλλο
«Dashboard» με τα 7 διαγράμματα χάνεται σιωπηλά. Το script το ξαναβάζει.

Πώς δουλεύει: χειρουργική ένθεση στο ZIP του .xlsx. ΔΕΝ ανοίγει το βιβλίο με
openpyxl, άρα δεν πειράζει τίποτε άλλο — τύποι, πίνακες, μορφοποίηση υπό όρους,
επικύρωση δεδομένων και τα checkboxes του φύλλου «Στατιστικά» μένουν ακέραια.

Χρήση:
    python3 rebuild_dashboard.py "ΠΥΡΚΑΓΙΕΣ 2026.xlsx"            # επιτόπου
    python3 rebuild_dashboard.py in.xlsx -o out.xlsx              # σε νέο αρχείο
    python3 rebuild_dashboard.py in.xlsx --rows 3000              # άλλο όριο γραμμών

Το πρότυπο του φύλλου βρίσκεται στο dashboard_template/ δίπλα στο script.
"""

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

TPL = Path(__file__).resolve().parent / "dashboard_template"
SHEET_NAME = "Dashboard"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

# Το πρότυπο γράφτηκε με σταθερά παράθυρα 500 γραμμών (Ελλάδα) και 300 (Ευρώπη).
TPL_GR_ROWS, TPL_EU_ROWS = 500, 300


def widen(xml: str, gr_rows: int, eu_rows: int) -> str:
    """Διευρύνει τα σταθερά εύρη ώστε να καλύπτουν όλα τα δεδομένα."""
    xml = re.sub(r"('Ελλάδα'!\$[A-Z]{1,3}\$\d+:\$[A-Z]{1,3}\$)%d\b" % TPL_GR_ROWS,
                 lambda m: m.group(1) + str(gr_rows), xml)
    xml = re.sub(r"('Ευρώπη'!\$[A-Z]{1,3}\$\d+:\$[A-Z]{1,3}\$)%d\b" % TPL_EU_ROWS,
                 lambda m: m.group(1) + str(eu_rows), xml)
    return xml


def dashboard_parts(entries: dict) -> set:
    """Εντοπίζει τα μέρη που ανήκουν σε τυχόν υπάρχον φύλλο Dashboard."""
    wb = entries.get("xl/workbook.xml", b"").decode("utf-8")
    m = re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"' % SHEET_NAME, wb)
    if not m:
        return set()
    rid = m.group(1)
    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    t = re.search(r'<Relationship Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
    if not t:
        return set()
    sheet_path = "xl/" + t.group(1).lstrip("/")
    doomed = {sheet_path}
    srels_path = f"xl/worksheets/_rels/{Path(sheet_path).name}.rels"
    srels = entries.get(srels_path, b"").decode("utf-8")
    doomed.add(srels_path)
    dm = re.search(r'Target="\.\./(drawings/[^"]+)"', srels)
    if dm:
        draw = "xl/" + dm.group(1)
        doomed.add(draw)
        drels_path = f"xl/drawings/_rels/{Path(draw).name}.rels"
        doomed.add(drels_path)
        for c in re.findall(r'Target="\.\./(charts/[^"]+)"',
                            entries.get(drels_path, b"").decode("utf-8")):
            doomed.add("xl/" + c)
    return {d for d in doomed if d in entries}


def free_name(entries, pattern, start=1):
    i = start
    while pattern % i in entries:
        i += 1
    return i


def rebuild(src: Path, dst: Path, gr_rows: int, eu_rows: int) -> dict:
    with zipfile.ZipFile(src) as z:
        entries = {n: z.read(n) for n in z.namelist()}
        order = [n for n in z.namelist()]

    removed = dashboard_parts(entries)
    for p in removed:
        entries.pop(p, None)
    order = [n for n in order if n not in removed]

    # --- διάλεξε ελεύθερα ονόματα μερών ---
    si = free_name(entries, "xl/worksheets/sheet%d.xml")
    di = free_name(entries, "xl/drawings/drawing%d.xml")
    c0 = free_name(entries, "xl/charts/chart%d.xml")
    sheet_part = f"xl/worksheets/sheet{si}.xml"
    draw_part = f"xl/drawings/drawing{di}.xml"
    chart_parts = [f"xl/charts/chart{c0 + k}.xml" for k in range(7)]

    # --- workbook.xml.rels: νέο rId για το φύλλο ---
    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    used = {int(x) for x in re.findall(r'Id="rId(\d+)"', rels)}
    rid = f"rId{max(used) + 1 if used else 1}"
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="{rid}" Type="{NS_R}/worksheet" '
                        f'Target="worksheets/sheet{si}.xml"/></Relationships>')
    entries["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # --- workbook.xml: καταχώριση του φύλλου στο τέλος ---
    wb = entries["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r'<sheet name="%s"[^>]*/>' % SHEET_NAME, "", wb)
    sid = max([int(x) for x in re.findall(r'sheetId="(\d+)"', wb)] or [0]) + 1
    wb = wb.replace("</sheets>",
                    f'<sheet name="{SHEET_NAME}" sheetId="{sid}" r:id="{rid}"/></sheets>')
    entries["xl/workbook.xml"] = wb.encode("utf-8")

    # --- τα μέρη του Dashboard από το πρότυπο ---
    sheet_xml = widen((TPL / "sheet.xml").read_text("utf-8"), gr_rows, eu_rows)
    entries[sheet_part] = sheet_xml.encode("utf-8")

    entries[f"xl/worksheets/_rels/sheet{si}.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{NS_R}/drawing" '
        f'Target="../drawings/drawing{di}.xml"/></Relationships>').encode("utf-8")

    entries[draw_part] = (TPL / "drawing.xml").read_bytes()
    drels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for k, cp in enumerate(chart_parts, 1):
        drels.append(f'<Relationship Id="rId{k}" Type="{NS_R}/chart" '
                     f'Target="../charts/{Path(cp).name}"/>')
    drels.append("</Relationships>")
    entries[f"xl/drawings/_rels/drawing{di}.xml.rels"] = "".join(drels).encode("utf-8")

    for k, cp in enumerate(chart_parts, 1):
        entries[cp] = widen((TPL / f"chart{k}.xml").read_text("utf-8"),
                            gr_rows, eu_rows).encode("utf-8")

    # --- [Content_Types].xml ---
    ct = entries["[Content_Types].xml"].decode("utf-8")
    for p in removed:
        ct = ct.replace(f'<Override PartName="/{p}" ContentType="{CT_SHEET}"/>', "")
        ct = ct.replace(f'<Override PartName="/{p}" ContentType="{CT_DRAWING}"/>', "")
        ct = ct.replace(f'<Override PartName="/{p}" ContentType="{CT_CHART}"/>', "")
    adds = [f'<Override PartName="/{sheet_part}" ContentType="{CT_SHEET}"/>',
            f'<Override PartName="/{draw_part}" ContentType="{CT_DRAWING}"/>']
    adds += [f'<Override PartName="/{c}" ContentType="{CT_CHART}"/>' for c in chart_parts]
    ct = ct.replace("</Types>", "".join(adds) + "</Types>")
    entries["[Content_Types].xml"] = ct.encode("utf-8")

    # --- γράψιμο ---
    new = [n for n in entries if n not in order]
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order + new:
            z.writestr(n, entries[n])
    shutil.move(str(tmp), str(dst))
    return {"removed": sorted(removed), "sheet": sheet_part,
            "charts": len(chart_parts), "gr_rows": gr_rows, "eu_rows": eu_rows}


def main():
    ap = argparse.ArgumentParser(description="Αποκατάσταση του φύλλου Dashboard")
    ap.add_argument("workbook", type=Path)
    ap.add_argument("-o", "--output", type=Path, help="default: επιτόπου")
    ap.add_argument("--rows", type=int, default=2000, help="όριο γραμμών «Ελλάδα» (default 2000)")
    ap.add_argument("--eu-rows", type=int, default=1000, help="όριο γραμμών «Ευρώπη» (default 1000)")
    a = ap.parse_args()
    if not a.workbook.exists():
        sys.exit(f"Δεν βρέθηκε: {a.workbook}")
    for f in ["sheet.xml", "drawing.xml"] + [f"chart{i}.xml" for i in range(1, 8)]:
        if not (TPL / f).exists():
            sys.exit(f"Λείπει το πρότυπο: {TPL / f}")
    out = a.output or a.workbook
    r = rebuild(a.workbook, out, a.rows, a.eu_rows)
    was = "αντικαταστάθηκε" if r["removed"] else "προστέθηκε"
    print(f"OK: το φύλλο «{SHEET_NAME}» {was} ({r['charts']} γραφήματα) -> {out}")
    print(f"    εύρη: Ελλάδα $2:${r['gr_rows']} · Ευρώπη $2:${r['eu_rows']}")


if __name__ == "__main__":
    main()
