"""Eingabevorlage im Zielformat: DATEV-BWA-Kostenstellenblätter (Form 01).

Aufruf: python -m prognosemodell.vorlagen [--ausgabe PFAD]

Schreibt vorlagen/Kostenstellen_BWA_Vorlage.xlsx mit erfundenen Werten:
ein Summenblatt "Alle Objekte" und je Kostenstelle ein Blatt im selben Layout.
Spalten wie in der Planungsreferenz der Kanzlei (Projektplan, Abschnitt 18):
F, G Vorjahre, H–S Monate des Basisjahrs, T Jahr Basisjahr, U–AN Planjahre.
Die Planspalten bleiben leer; sie füllt später das Modell.
"""

import argparse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .modelle import PARAMETER

# (BWA-Nr., Bezeichnung) in der Reihenfolge der BWA Form 01; leere Bezeichnung = Leerzeile
BWA_ZEILEN = [
    (1010, ""), (1020, "Umsatzerlöse"), (1050, ""), (1051, "Gesamtleistung"),
    (1052, ""), (1070, ""), (1080, "Rohertrag"), (1081, ""), (1090, "So. betr. Erlöse"),
    (1091, ""), (1092, "Betriebl. Rohertrag"), (1093, ""), (1094, "Kostenarten:"),
    (1100, "Personalkosten"), (1120, "Raumkosten"), (1140, "Betriebl. Steuern"),
    (1150, "Versich./Beiträge"), (1180, "Kfz-Kosten (o. St.)"),
    (1200, "Werbe-/Reisekosten"), (1220, "Kosten Warenabgabe"), (1240, "Abschreibungen"),
    (1250, "Reparatur/Instandh."), (1260, "Sonstige Kosten"), (1280, "Gesamtkosten"),
    (1290, ""), (1300, "Betriebsergebnis"), (1301, ""), (1310, "Zinsaufwand"),
    (1312, "Sonst. neutr. Aufw"), (1320, "Neutraler Aufwand"), (1321, ""),
    (1322, "Zinserträge"), (1323, "Sonst. neutr. Ertr"), (1330, "Neutraler Ertrag"),
    (1331, ""), (1342, ""), (1345, "Ergebnis vor Steuern"), (1350, ""),
    (1351, "außerordentlicher Ertrag"), (1352, "außerordentlicher Aufwand"),
    (1353, "Außerordentliches Ergebnis"), (1354, ""), (1355, "Steuern Eink.u.Ertr"),
    (1360, ""), (1380, "Vorläufiges Ergebnis"),
]
KOSTENARTEN = (1100, 1120, 1140, 1150, 1180, 1200, 1220, 1240, 1250, 1260)

# Layout, wie es einlesen.py erwartet (B2/C2 Kopf, Zeile 4 Spaltenköpfe, ab Zeile 6 BWA)
ERSTE_ZEILE = 6
SPALTE_VORJAHRE = 6     # F: Basisjahr − 2, G: Basisjahr − 1
SPALTE_MONATE = 8       # H–S
SPALTE_JAHR = 20        # T
SPALTE_PLAN = 21        # U–AN

# Erfundene Jahreswerte im Basisjahr; KSt 1 hat Miete und Erhaltung des Testobjekts
BEISPIELE = [
    ("KSt 1", "Musterstraße 1", {1020: 60_000, 1150: 1_800, 1140: 1_200,
                                  1240: 16_000, 1250: 8_000, 1260: 600, 1310: 4_000}),
    ("KSt 2", "Beispielweg 7", {1020: 120_000, 1090: 1_500, 1120: 2_400, 1140: 2_600,
                                 1150: 3_200, 1240: 30_000, 1250: 14_000, 1260: 1_100,
                                 1310: 9_000, 1322: 300}),
]
VORJAHRESFAKTOR = (0.96, 0.98)  # Basisjahr − 2, − 1
GELB = PatternFill("solid", fgColor="FFF2CC")


def _summen(werte: dict) -> dict:
    """Summenzeilen der BWA aus den Einzelzeilen."""
    w = {nr: werte.get(nr, 0.0) for nr, _ in BWA_ZEILEN}
    w[1051] = w[1080] = w[1020]
    w[1092] = w[1080] + w[1090]
    w[1280] = sum(w[nr] for nr in KOSTENARTEN)
    w[1300] = w[1092] - w[1280]
    w[1320] = w[1310] + w[1312]
    w[1330] = w[1322] + w[1323]
    w[1345] = w[1300] - w[1320] + w[1330]
    w[1353] = w[1345] + w[1351] - w[1352]
    w[1380] = w[1353] - w[1355]
    return w


def kostenstellenblatt(ws, kst: str, name: str, werte: dict, basisjahr: int,
                       planjahre: int = 20) -> None:
    """Ein Blatt im BWA-Layout; Monate = Jahreswert / 12, Planspalten leer."""
    ws["B2"], ws["C2"] = kst, name
    ws["B2"].font = ws["C2"].font = Font(bold=True)
    ws["B4"], ws["C4"] = "Nr.", "Bezeichnung kurz"
    ws.cell(row=4, column=SPALTE_VORJAHRE, value=f"Jahr {basisjahr - 2}")
    ws.cell(row=4, column=SPALTE_VORJAHRE + 1, value=f"Jahr {basisjahr - 1}")
    for m in range(12):
        zelle = ws.cell(row=4, column=SPALTE_MONATE + m, value=datetime(basisjahr, m + 1, 1))
        zelle.number_format = "MMM YY"
    ws.cell(row=4, column=SPALTE_JAHR, value=f"Jahr {basisjahr}")
    for i in range(planjahre):
        jahr = basisjahr + 1 + i
        zelle = ws.cell(row=4, column=SPALTE_PLAN + i,
                        value=f"Plan {jahr}" if i == 0 else jahr)
        zelle.fill = GELB
    for spalte in range(2, SPALTE_PLAN + planjahre):
        ws.cell(row=4, column=spalte).font = Font(bold=True)

    summe = _summen(werte)
    for i, (nr, bezeichnung) in enumerate(BWA_ZEILEN):
        zeile = ERSTE_ZEILE + i
        ws.cell(row=zeile, column=2, value=nr)
        if not bezeichnung:
            continue
        ws.cell(row=zeile, column=3, value=bezeichnung)
        wert = summe[nr]
        if wert == 0:
            continue
        for k, faktor in enumerate(VORJAHRESFAKTOR):
            ws.cell(row=zeile, column=SPALTE_VORJAHRE + k, value=round(wert * faktor, 2))
        for m in range(12):
            ws.cell(row=zeile, column=SPALTE_MONATE + m, value=round(wert / 12, 2))
        ws.cell(row=zeile, column=SPALTE_JAHR, value=wert)
    for spalte in range(SPALTE_VORJAHRE, SPALTE_PLAN + planjahre):
        for zeile in range(ERSTE_ZEILE, ERSTE_ZEILE + len(BWA_ZEILEN)):
            ws.cell(row=zeile, column=spalte).number_format = "#,##0.00"
        ws.column_dimensions[get_column_letter(spalte)].width = 12
    ws.column_dimensions["C"].width = 26
    ws.freeze_panes = "D5"


def erstelle_vorlage(beispiele=BEISPIELE, basisjahr: int = None) -> Workbook:
    if basisjahr is None:
        basisjahr = next(p.wert for p in PARAMETER if p.name == "par_Basisjahr")
    wb = Workbook()
    wb.remove(wb.active)
    gesamt = {}
    for _, _, werte in beispiele:
        for nr, wert in werte.items():
            gesamt[nr] = gesamt.get(nr, 0.0) + wert
    kostenstellenblatt(wb.create_sheet("Alle Objekte"), "KSt", "Alle Objekte", gesamt,
                       basisjahr)
    for kst, name, werte in beispiele:
        kostenstellenblatt(wb.create_sheet(kst), kst, name, werte, basisjahr)
    return wb


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ausgabe", default="vorlagen/Kostenstellen_BWA_Vorlage.xlsx")
    args = ap.parse_args()
    ziel = Path(args.ausgabe)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    erstelle_vorlage().save(ziel)
    print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
