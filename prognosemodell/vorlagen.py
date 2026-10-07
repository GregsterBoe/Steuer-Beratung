"""Eingabevorlagen im Zielformat: DATEV-BWA-Kostenstellenblätter (Form 01) und
Anlagenverzeichnis (DATEV Inventarübersicht).

Aufruf: python -m prognosemodell.vorlagen [--ausgabe PFAD] [--inventar PFAD]

Schreibt vorlagen/Kostenstellen_BWA_Vorlage.xlsx mit erfundenen Werten:
ein Summenblatt "Alle Objekte" und je Kostenstelle ein Blatt im selben Layout.
Spalten wie in der Planungsreferenz der Kanzlei (Projektplan, Abschnitt 18):
F, G Vorjahre, H–S Monate des Basisjahrs, T Jahr Basisjahr, U–AN Planjahre.
Die Planspalten bleiben leer; sie füllt später das Modell. Unter der BWA steht die
Aufschlüsselung der Abschreibungen je Anlagengruppe (Projektplan, Abschnitt 21).

Dazu vorlagen/Inventar_Vorlage.xlsx: Anlagenverzeichnis im Spaltenformat des
DATEV-Exports, ebenfalls erfunden, KOST1 verweist auf die Kostenstelle.
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
# Aufschlüsselung der Abschreibungen je Kostenstelle, Stand Ende Basisjahr − 1:
# (Beschriftung Buchwert, Beschriftung AfA, Buchwert, Jahres-AfA); wie im Muster der
# Kanzlei dürfen die Beschriftungen abweichen („Wohnbau 7“ zu „Wohnbau“)
AUFSCHLUESSELUNG = {
    "KSt 1": [("Wohnbau 1", "Wohnbau 1", 496_000, 16_000)],
    "KSt 2": [("Wohnbau 7", "Wohnbau", 876_000, 24_000),
              ("Außenanlagen 7", "Außenanlagen", 15_000, 6_000),
              ("Anbau im Bau", "Anbau im Bau", 50_000, 0),
              ("sonstige", "sonstige", None, 0)],
}
GELB = PatternFill("solid", fgColor="FFF2CC")
FMT_BWA = "#,##0.00"


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


def bwa_kopf(ws, kst: str, name: str, basisjahr: int, planjahre: int = 20) -> dict:
    """Kopf und Zeilenbeschriftung eines BWA-Blatts; liefert BWA-Nr. -> Blattzeile.

    Gemeinsam für die Vorlage und die BWA-Ausgabe der Mappe (bwa.py).
    """
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

    zeilen = {}
    for i, (nr, bezeichnung) in enumerate(BWA_ZEILEN):
        zeile = ERSTE_ZEILE + i
        zeilen[nr] = zeile
        ws.cell(row=zeile, column=2, value=nr)
        if bezeichnung:
            ws.cell(row=zeile, column=3, value=bezeichnung)
    for spalte in range(SPALTE_VORJAHRE, SPALTE_PLAN + planjahre):
        for zeile in zeilen.values():
            ws.cell(row=zeile, column=spalte).number_format = FMT_BWA
        ws.column_dimensions[get_column_letter(spalte)].width = 12
    ws.column_dimensions["C"].width = 26
    ws.freeze_panes = "D5"
    return zeilen


def aufschluesselung(ws, gruppen: list, erste_zeile: int) -> None:
    """Blöcke „Buchwert, JE“, „Abschreibungen JW“, „Abschreibungen MW“ in Spalte C,
    Werte in der Spalte Basisjahr − 1 (G), Summenzeilen mit Wert, Zwischensumme ohne Text."""
    sp = SPALTE_VORJAHRE + 1
    zeile = erste_zeile
    ws.cell(row=zeile, column=3, value="Buchwert, JE").font = Font(bold=True)
    ws.cell(row=zeile, column=sp, value=sum(g[2] or 0 for g in gruppen))
    for bw_text, _, bw, _ in gruppen:
        zeile += 1
        ws.cell(row=zeile, column=3, value=bw_text)
        ws.cell(row=zeile, column=sp, value=bw)
    for kopf, faktor in (("Abschreibungen JW", 1), ("Abschreibungen MW", 1 / 12)):
        zeile += 2
        ws.cell(row=zeile, column=3, value=kopf).font = Font(bold=True)
        ws.cell(row=zeile, column=sp, value=sum(g[3] for g in gruppen) * faktor)
        if faktor == 1:
            zeile += 1   # Zwischensumme ohne Beschriftung, wie im Muster
            ws.cell(row=zeile, column=sp, value=sum(g[3] for g in gruppen))
        for _, afa_text, _, afa in gruppen:
            zeile += 1
            ws.cell(row=zeile, column=3, value=afa_text)
            ws.cell(row=zeile, column=sp, value=round(afa * faktor, 2))


def kostenstellenblatt(ws, kst: str, name: str, werte: dict, basisjahr: int,
                       planjahre: int = 20, gruppen: list = None) -> None:
    """Ein Blatt im BWA-Layout; Monate = Jahreswert / 12, Planspalten leer."""
    zeilen = bwa_kopf(ws, kst, name, basisjahr, planjahre)
    if gruppen:
        aufschluesselung(ws, gruppen, max(zeilen.values()) + 3)
    summe = _summen(werte)
    for nr, bezeichnung in BWA_ZEILEN:
        wert = summe[nr]
        if not bezeichnung or wert == 0:
            continue
        zeile = zeilen[nr]
        for k, faktor in enumerate(VORJAHRESFAKTOR):
            ws.cell(row=zeile, column=SPALTE_VORJAHRE + k, value=round(wert * faktor, 2))
        for m in range(12):
            ws.cell(row=zeile, column=SPALTE_MONATE + m, value=round(wert / 12, 2))
        ws.cell(row=zeile, column=SPALTE_JAHR, value=wert)


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
        kostenstellenblatt(wb.create_sheet(kst), kst, name, werte, basisjahr,
                           gruppen=AUFSCHLUESSELUNG.get(kst))
    return wb


# --- Anlagenverzeichnis im Format des DATEV-Exports „Inventarübersicht“ ---

INVENTAR_KOPF = [
    "Konto", "Inventar", "Inventarbezeichnung", "AHK-Datum", "AHK Wj-Ende", "Buchw. Wj-Ende",
    "N-AfA Wj-Ende", "S-Abschr. Wj-Ende", "ND", "AfA-Art", "AfA-%", "KOST1", "KOST2", "Filiale",
    "Lieferanten-Nr.", "ANLAG-Lieferant", "AHK Wj-Beginn", "Buchw. Wj-Beginn", "N-AfA Wj-Beginn",
    "S-Abschr. Wj-Beginn", "S-Abschr.Art", "S-Abschr.%", "Restbegünst.", "S-Abschr.Verteil.",
    "Abgang", "Lebenslaufakte", "Bestelldatum", "Erl. AfA-Art", "Herkunftsart", "WKN/ISIN",
    "Erfassungsart",
]
# Erfunden, Stand Ende Basisjahr − 1 (2025). KSt 1 = Testobjekt: Gebäude 800.000 zu 2 %,
# G+B 200.000, Kauf 2007, Restbuchwert Ende 2026 480.000. KSt 2 mit Außenanlage, die 2028
# ausläuft, und einer Anlage im Bau. Dazu Fälle ohne Objekt, Finanzanlage, Abgang.
# (Konto, Inventar, Bezeichnung, AHK-Datum, AHK, Buchwert Ende, N-AfA Beginn, N-AfA Ende,
#  ND, AfA-Art, AfA-%, KOST1, Abgang)
INVENTAR_BEISPIELE = [
    (235, "100001", "Grund und Boden Musterstraße 1", "15.03.2007", 200_000, 200_000, None,
     None, "", "Keine AfA", None, 1, None),
    (300, "300001", "Wohngebäude Musterstraße 1", "15.03.2007", 800_000, 496_000, 288_000,
     304_000, "50/00", "Lin.Geb.12", 2, 1, None),
    (235, "100002", "Grund und Boden Beispielweg 7", "01.07.2012", 400_000, 400_000, None,
     None, "", "Keine AfA", None, 2, None),
    (300, "300002", "Wohngebäude Beispielweg 7", "01.07.2012", 1_200_000, 876_000, 300_000,
     324_000, "50/00", "Lin.Geb.12", 2, 2, None),
    (310, "310002", "Außenanlagen Beispielweg 7", "01.07.2018", 60_000, 15_000, 39_000,
     45_000, "10/00", "Linear", 10, 2, None),
    (750, "750002", "Anbau im Bau Beispielweg 7", "31.12.2025", 50_000, 50_000, None, None,
     "", "Anlag./Bau", None, 2, None),
    (235, "100009", "Grund und Boden ohne Kostenstelle", "01.01.2015", 90_000, 90_000, None,
     None, "", "Keine AfA", None, None, None),
    (300, "300009", "Wohngebäude Kostenstelle ohne Blatt", "01.01.2015", 300_000, 234_000,
     60_000, 66_000, "50/00", "Lin.Geb.12", 2, 9, None),
    (690, "690003", "Büroausstattung degressiv", "01.01.2024", 1_000, 800, 0, 200, "08/00",
     "Geom.degr.", 25, None, None),
    (520, "520001", "Pkw abgegangen", "01.01.2020", 30_000, 0, 30_000, 30_000, "06/00",
     "Linear", 16.67, 2, "15.06.2025"),
    (910, "910001", "Wertpapiere", "01.01.2018", 25_000, 25_000, None, None, "",
     "Finanzanl.", None, None, None),
]


def erstelle_inventar_vorlage(beispiele=INVENTAR_BEISPIELE) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventarübersicht"
    for spalte, text in enumerate(INVENTAR_KOPF, start=1):
        ws.cell(row=1, column=spalte, value=text).font = Font(bold=True)
    spalte = {text: i for i, text in enumerate(INVENTAR_KOPF, start=1)}
    for zeile, (konto, nr, bez, datum, ahk, bw, afa_beginn, afa_ende, nd, art, satz, kost,
                abgang) in enumerate(beispiele, start=2):
        werte = {"Konto": konto, "Inventar": nr, "Inventarbezeichnung": bez,
                 "AHK-Datum": datum, "AHK Wj-Ende": 0 if abgang else ahk,
                 "Buchw. Wj-Ende": bw, "N-AfA Wj-Ende": afa_ende or 0, "ND": nd,
                 "AfA-Art": art, "AfA-%": satz, "KOST1": kost if kost is not None else "",
                 "AHK Wj-Beginn": ahk, "Buchw. Wj-Beginn": bw + (afa_ende or 0) - (afa_beginn or 0),
                 "N-AfA Wj-Beginn": afa_beginn, "Abgang": abgang}
        for text, wert in werte.items():
            if wert is not None:
                ws.cell(row=zeile, column=spalte[text], value=wert)
    for i in range(1, len(INVENTAR_KOPF) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.column_dimensions["C"].width = 32
    return wb


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ausgabe", default="vorlagen/Kostenstellen_BWA_Vorlage.xlsx")
    ap.add_argument("--inventar", default="vorlagen/Inventar_Vorlage.xlsx")
    args = ap.parse_args()
    for ziel, mappe in ((Path(args.ausgabe), erstelle_vorlage()),
                        (Path(args.inventar), erstelle_inventar_vorlage())):
        ziel.parent.mkdir(parents=True, exist_ok=True)
        mappe.save(ziel)
        print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
