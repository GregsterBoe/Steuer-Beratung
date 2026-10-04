"""Vorlagen für die beiden Eingabedateien, mit Beispieldaten.

Aufruf: python -m prognosemodell.vorlagen [--ordner vorlagen]

Das Layout kommt aus einlesen.py; was hier geschrieben wird, liest einlesen.py
wieder ein. Objekt OBJ-001 entspricht dem Testobjekt aus testdaten.py.
"""

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from .einlesen import (BLATT_STAMMDATEN, BLATT_ZUORDNUNG, KATEGORIE_AUSGABEN,
                       KATEGORIE_EINNAHMEN, KATEGORIE_ERHALTUNG, KATEGORIE_IGNORIEREN,
                       KATEGORIE_MIETE, KATEGORIEN, KST_KOPF_BEZEICHNUNG,
                       KST_KOPF_JAHR, KST_KOPF_KOSTENSTELLE, KST_SPALTE_BETRAG,
                       KST_SPALTE_KONTO, KST_SPALTE_TEXT, KST_TITEL, STAMM_FELDER,
                       STAMM_KOPFZEILE, ZUORD_FELDER, ZUORD_KOPFZEILE)
from .mappe import FILL_EINGABE, FONT_TITEL, _kopf
from .modelle import FMT_EURO

DATEI_KOSTENSTELLEN = "Kostenstellen_Vorlage.xlsx"
DATEI_STAMMDATEN = "Stammdaten_Vorlage.xlsx"
GESCHAEFTSJAHR = 2026
FONT_FETT = Font(bold=True)

# Kontenzuordnung: (von, bis, Kategorie, Bemerkung); Kontonummern sind Beispiele
ZUORDNUNG = [
    (4100, 4199, KATEGORIE_MIETE, "Mieterlöse Wohnen und Gewerbe"),
    (4200, 4299, KATEGORIE_EINNAHMEN, "Umlagen Nebenkosten"),
    (6220, 6229, KATEGORIE_IGNORIEREN, "Abschreibungen: rechnet das Modell selbst"),
    (6300, 6309, KATEGORIE_AUSGABEN, "Hausverwaltung"),
    (6325, 6329, KATEGORIE_AUSGABEN, "Strom, Gas, Wasser"),
    (6335, 6339, KATEGORIE_ERHALTUNG, "Instandhaltung Gebäude"),
    (6350, 6359, KATEGORIE_AUSGABEN, "Grundsteuer"),
    (6400, 6409, KATEGORIE_AUSGABEN, "Versicherungen"),
    (6460, 6469, KATEGORIE_ERHALTUNG, "Reparaturen"),
    (7300, 7399, KATEGORIE_IGNORIEREN, "Zinsen: Finanzierung kommt in Stufe 2"),
]

# Kostenstellen: (Nummer, Bezeichnung, [(Konto, Text, Betrag)])
KOSTENSTELLEN = [
    (1001, "Wohnanlage Musterstraße 1", [
        (4105, "Mieterlöse Wohnraum", 48_000),
        (4106, "Mieterlöse Gewerbe", 12_000),
        (4210, "Umlagen Nebenkosten", 9_500),
        (6222, "Abschreibung Gebäude", 20_000),
        (6300, "Hausverwaltung", 3_600),
        (6325, "Strom, Gas, Wasser", 6_800),
        (6335, "Instandhaltung Gebäude", 5_000),
        (6350, "Grundsteuer", 2_100),
        (6400, "Versicherungen", 1_400),
        (6460, "Reparaturen", 3_000),
        (7310, "Zinsaufwand Darlehen", 11_000),
    ]),
    (1002, "Gewerbeeinheit Hafenweg 7", [
        (4106, "Mieterlöse Gewerbe", 95_000),
        (4210, "Umlagen Nebenkosten", 14_000),
        (6222, "Abschreibung Gebäude", 36_000),
        (6335, "Instandhaltung Gebäude", 12_000),
        (6350, "Grundsteuer", 4_300),
        (6400, "Versicherungen", 2_900),
    ]),
    (1003, "Mehrfamilienhaus Lindenallee 12", [
        (4105, "Mieterlöse Wohnraum", 52_000),
        (4210, "Umlagen Nebenkosten", 11_200),
        (6222, "Abschreibung Gebäude", 12_000),
        (6300, "Hausverwaltung", 2_400),
        (6460, "Reparaturen", 9_500),
        (6350, "Grundsteuer", 1_800),
    ]),
]

# Stammdaten je Objekt; Restbuchwert = Ende Geschäftsjahr
STAMMDATEN = [
    dict(objekt_id="OBJ-001", kostenstelle=1001, name="Testobjekt", ak_gebaeude=800_000,
         ak_gub=200_000, kaufjahr=2007, afa_satz=0.025, restbuchwert=400_000,
         verkehrswert=1_400_000, vk_quote_gebaeude=0.5),
    dict(objekt_id="OBJ-002", kostenstelle=1002, name=None, ak_gebaeude=1_200_000,
         ak_gub=300_000, kaufjahr=2015, afa_satz=0.03, restbuchwert=768_000,
         verkehrswert=1_900_000, vk_quote_gebaeude=0.6),
    dict(objekt_id="OBJ-003", kostenstelle=1003, name=None, ak_gebaeude=600_000,
         ak_gub=150_000, kaufjahr=1998, afa_satz=0.02, restbuchwert=252_000,
         verkehrswert=1_250_000, vk_quote_gebaeude=None),
]

ANLEITUNG_KOSTENSTELLEN = [
    "So muss die Kostenstellendatei aussehen",
    "",
    "Ein Blatt je Kostenstelle, also je Immobilie. Der Blattname ist frei.",
    f"Im Kopf (Zeilen 1 bis 10) stehen in Spalte A die Beschriftungen "
    f"'{KST_KOPF_KOSTENSTELLE}', '{KST_KOPF_BEZEICHNUNG}' und '{KST_KOPF_JAHR}', "
    "der Wert jeweils rechts daneben in Spalte B.",
    f"Darunter die Kontentabelle. Ihre Kopfzeile beginnt in Spalte A mit "
    f"'{KST_SPALTE_KONTO}' und enthält '{KST_SPALTE_TEXT}' und '{KST_SPALTE_BETRAG}'.",
    "Je Konto eine Zeile mit dem Jahresbetrag des Geschäftsjahres. Erlöse und Aufwand "
    "beide positiv eintragen; ob ein Konto Einnahme oder Ausgabe ist, sagt die "
    "Kontenzuordnung in der Stammdatendatei.",
    "Zeilen ohne Kontonummer (Leerzeilen, Summen) werden übergangen.",
    "Blätter ohne die Beschriftung Kostenstelle, wie dieses, werden übergangen.",
    "Alle Blätter müssen dasselbe Geschäftsjahr zeigen; es wird das Basisjahr des Modells.",
    "",
    "Weicht die echte Kanzlei-Excel davon ab, werden nur die Konstanten oben in "
    "prognosemodell/einlesen.py angepasst.",
]

ANLEITUNG_STAMMDATEN = [
    "So muss die Stammdatendatei aussehen",
    "",
    f"Blatt '{BLATT_STAMMDATEN}': eine Zeile je Immobilie mit den steuerlichen Stammdaten, "
    "meist aus dem Anlageverzeichnis. Gelbe Spalten sind Pflicht.",
    "Die Spalte Kostenstelle verbindet die Zeile mit dem Blatt in der Kostenstellendatei; "
    "Miete, Erhaltung und die weiteren Einnahmen und Ausgaben kommen von dort.",
    "Objektname leer: die Bezeichnung aus dem Kostenstellenblatt wird übernommen.",
    "Restbuchwert Gebäude zum Ende des Geschäftsjahres der Kostenstellen (= Basisjahr).",
    "Verkehrswert und Verkehrswertanteil Gebäude sind optional.",
    "Die Überschriften müssen genau so bleiben; die Reihenfolge der Spalten ist frei.",
    "",
    f"Blatt '{BLATT_ZUORDNUNG}': ordnet jedes Konto der Kostenstellen einer Kategorie zu: "
    + ", ".join(KATEGORIEN) + ".",
    "Ein Konto mit Betrag, das keiner Zeile zugeordnet ist, bricht das Einlesen mit "
    "Meldung ab; so geht kein Betrag verloren.",
    "Weitere Einnahmen und Ausgaben stehen im Objektblatt, fließen aber noch nicht ins "
    "Ergebnis (Projektplan Abschnitt 8).",
]


def _anleitung(ws, zeilen: list) -> None:
    ws.title = "Anleitung"
    for i, text in enumerate(zeilen, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws["A1"].font = FONT_TITEL
    ws.column_dimensions["A"].width = 110


def kostenstellen_vorlage(geschaeftsjahr: int = GESCHAEFTSJAHR) -> Workbook:
    wb = Workbook()
    _anleitung(wb.active, ANLEITUNG_KOSTENSTELLEN)
    for nummer, bezeichnung, konten in KOSTENSTELLEN:
        ws = wb.create_sheet(f"KST {nummer}")
        ws["A1"] = KST_TITEL
        ws["A1"].font = FONT_TITEL
        for zeile, (beschriftung, wert) in enumerate(
                [(KST_KOPF_KOSTENSTELLE, nummer), (KST_KOPF_BEZEICHNUNG, bezeichnung),
                 (KST_KOPF_JAHR, geschaeftsjahr)], start=3):
            ws.cell(row=zeile, column=1, value=beschriftung).font = FONT_FETT
            ws.cell(row=zeile, column=2, value=wert).alignment = Alignment(horizontal="left")
        _kopf(ws, 7, [KST_SPALTE_KONTO, KST_SPALTE_TEXT, KST_SPALTE_BETRAG])
        for zeile, (konto, text, betrag) in enumerate(konten, start=8):
            ws.cell(row=zeile, column=1, value=konto)
            ws.cell(row=zeile, column=2, value=text)
            ws.cell(row=zeile, column=3, value=betrag).number_format = FMT_EURO
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 34
        ws.column_dimensions["C"].width = 20
        ws.freeze_panes = "A8"
    return wb


def _tabelle(ws, titel: str, hinweis: str, kopfzeile: int, felder, saetze: list,
             leerzeilen: int) -> None:
    ws["A1"] = titel
    ws["A1"].font = FONT_TITEL
    ws["A2"] = hinweis
    _kopf(ws, kopfzeile, [f.ueberschrift for f in felder])
    ws.row_dimensions[kopfzeile].height = 45
    erste, letzte = kopfzeile + 1, kopfzeile + len(saetze) + leerzeilen
    for i, f in enumerate(felder, start=1):
        bst = ws.cell(row=kopfzeile, column=i).column_letter
        ws.column_dimensions[bst].width = f.breite
        for zeile in range(erste, letzte + 1):
            c = ws.cell(row=zeile, column=i)
            c.number_format = f.format
            if f.pflicht:
                c.fill = FILL_EINGABE
        if f.auswahl:
            dv = DataValidation(type="list", formula1='"' + ",".join(f.auswahl) + '"',
                                allow_blank=True, showErrorMessage=True)
            ws.add_data_validation(dv)
            dv.add(f"{bst}{erste}:{bst}{letzte}")
        for zeile, satz in enumerate(saetze, start=erste):
            ws.cell(row=zeile, column=i, value=satz.get(f.key))
    ws.freeze_panes = ws.cell(row=erste, column=2)


def stammdaten_vorlage() -> Workbook:
    wb = Workbook()
    _anleitung(wb.active, ANLEITUNG_STAMMDATEN)
    _tabelle(wb.create_sheet(BLATT_STAMMDATEN), "Stammdaten je Objekt",
             "gelb = Pflicht; Restbuchwert zum Ende des Geschäftsjahres der Kostenstellen",
             STAMM_KOPFZEILE, STAMM_FELDER, STAMMDATEN, 20)
    zuordnung = [dict(von=v, bis=b, kategorie=k, bemerkung=t) for v, b, k, t in ZUORDNUNG]
    _tabelle(wb.create_sheet(BLATT_ZUORDNUNG), "Kontenzuordnung",
             "jedes Konto der Kostenstellen mit Betrag braucht eine Zeile; Bereiche dürfen "
             "sich nicht überschneiden", ZUORD_KOPFZEILE, ZUORD_FELDER, zuordnung, 20)
    return wb


def erstelle_vorlagen(ordner: Path) -> tuple:
    ordner = Path(ordner)
    ordner.mkdir(parents=True, exist_ok=True)
    stamm, kst = ordner / DATEI_STAMMDATEN, ordner / DATEI_KOSTENSTELLEN
    stammdaten_vorlage().save(stamm)
    kostenstellen_vorlage().save(kst)
    return stamm, kst


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ordner", default="vorlagen", help="Zielordner, Standard vorlagen")
    for pfad in erstelle_vorlagen(Path(ap.parse_args().ordner)):
        print(f"geschrieben: {pfad}")


if __name__ == "__main__":
    main()
