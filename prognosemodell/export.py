"""Mandanten-Export: die gerechnete Mappe kompakt und nur mit Werten (Projektplan Abschnitt 30).

Aufruf: python -m prognosemodell.export MAPPE [--ausgabe PFAD]

Die Mappe muss durchgerechnet sein, also in Excel gespeichert. Ist sie es nicht (frisch
erzeugt und nie geöffnet), rechnet LibreOffice sie durch, sofern es installiert ist.

Statt aller Eingabe-, Rechen- und Kontrollblätter enthält der Export bis zu acht Blätter:
- Überblick: Empfehlung, Endvermögen je Option, Datenlage, Verkehrswert und Vermögen je Jahr
- Jahresrechnung: je Jahr Ergebnis und Steuer, Geldfluss und Vermögen, Plan und Halten
- Ergebnisbrücke: vom Ergebnis bei Halten je Verkauf und Kauf zum Ergebnis im Plan
  (entfällt ohne Verkauf und Kauf)
- BWA gesamt: Summen-BWA aller Objekte; mit „Vergleich mit Halten = ja“ auch die Differenz
  zu Halten
- BWA halten: Summen-BWA als Status quo, alle Objekte behalten (entfällt, wenn sie der BWA
  gesamt gleicht)
- BWA je Objekt: alle Kostenstellenblätter untereinander
- Verkauf und Kauf: Verkäufe, Neuobjekte und die Sicht je Vorgang (entfällt ohne beides)
- Annahmen: zentrale Annahmen, Daten je Objekt mit Herkunft, auffällige Prüfungen
Zeilen ohne Wert entfallen, ebenso Monatsspalten, Hilfszeilen und ausgeschaltete Vorgänge.
"""

import argparse
import datetime
import shutil
import subprocess
import sys
import tempfile
from copy import copy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .bwa import HALTEBLATT, SONDERBEREICH, SUMMENBLATT, SUMMENZEILEN
from .formeln import deckt_anlagen, hspalte, ispalte
from .mappe import (FILL_ANNAHME, FILL_EINGELESEN, FILL_KOPF, FONT_KOPF, OPTIONEN, _diagramm,
                    zentrale_annahmen)
from .modelle import (DARLEHEN_ANNAHMEN, FMT_EURO, KST_VERKNUEPFT, OBJEKT_EINGELESEN,
                      OBJEKT_FELDER, PARAMETER, PRUEFUNGEN, STATUS_AUS)
from .vorlagen import ERSTE_ZEILE, SPALTE_JAHR, SPALTE_PLAN, SPALTE_VORJAHRE

UEBERBLICK = "Überblick"
JAHRESRECHNUNG = "Jahresrechnung"
BRUECKE = "Ergebnisbrücke"
BWA_GESAMT = "BWA gesamt"
BWA_HALTEN = "BWA halten"
BWA_OBJEKTE = "BWA je Objekt"
VERKAUF_KAUF = "Verkauf und Kauf"
ANNAHMEN = "Annahmen"
INHALT = [
    (JAHRESRECHNUNG, "je Jahr Ergebnis und Steuer, Geldfluss und Vermögen; Plan und Halten"),
    (BRUECKE, "Wirkung jedes Verkaufs und Kaufs auf das Ergebnis vor Steuern, von Halten "
              "zum Plan"),
    (BWA_GESAMT, "Summen-BWA über alle Objekte: Ist-Jahre aus der Buchhaltung, dann die "
                 "Planjahre"),
    (BWA_HALTEN, "Summen-BWA als Status quo: alle Objekte behalten, ohne Verkäufe und "
                 "Neuobjekte"),
    (BWA_OBJEKTE, "BWA je Kostenstelle untereinander, mit Inhaltsverzeichnis"),
    (VERKAUF_KAUF, "geplante Verkäufe und Neuobjekte, Ergebnis und Herleitung je Vorgang"),
    (ANNAHMEN, "zentrale Annahmen, Daten je Objekt mit Herkunft, Plausibilitätsprüfung"),
]

FMT_BETRAG = '#,##0;[Red]-#,##0;"–"'
FMT_BWA = '#,##0;[Red]-#,##0;""'
FMT_DIFF = '[Color10]+#,##0;[Red]-#,##0;"–"'   # grün mehr, rot weniger als bei Halten
FONT_TITEL = Font(bold=True, size=14)
FONT_ABSCHNITT = Font(bold=True, size=12)
FONT_HINWEIS = Font(italic=True, color="595959")
FONT_FETT = Font(bold=True)
FONT_LINK = Font(color="0563C1", underline="single")
FILL_ABSCHNITT = PatternFill("solid", fgColor="D9E1F2")
RAHMEN_SUMME = Border(top=Side(style="thin"))
SCHWELLE = 0.005   # kleiner gilt als 0
BREITE_TEXT = 46
BREITE_ZAHL = 12


# --- Werte der Quellmappe ---

def _bereich(wb, name: str) -> list:
    """Werte eines benannten Bereichs, zeilen- oder spaltenweise als Liste."""
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    teile = ref.replace("$", "").split(":")
    zellen = wb[blatt][teile[0]:teile[-1]]
    return [c.value for zeile in zellen for c in zeile]


def _wert(wb, name: str):
    return _bereich(wb, name)[0]


def _zahl(wert) -> float:
    return wert if isinstance(wert, (int, float)) and not isinstance(wert, bool) else 0


def _ungleich_null(werte) -> bool:
    return any(abs(_zahl(w)) > SCHWELLE for w in werte)


# --- Gestaltung ---

def _titel(ws, titel: str, hinweis: str) -> int:
    ws["A1"] = titel
    ws["A1"].font = FONT_TITEL
    ws["A2"] = hinweis
    ws["A2"].font = FONT_HINWEIS
    return 4


def _abschnitt(ws, zeile: int, text: str, bis: int = 1) -> None:
    for spalte in range(1, bis + 1):
        ws.cell(row=zeile, column=spalte).fill = FILL_ABSCHNITT
    ws.cell(row=zeile, column=1, value=text).font = FONT_ABSCHNITT


def _kopfzeile(ws, zeile: int, texte: list, start: int = 1) -> None:
    for spalte, text in enumerate(texte, start=start):
        c = ws.cell(row=zeile, column=spalte, value=text)
        c.font, c.fill = FONT_KOPF, FILL_KOPF
        c.alignment = Alignment(horizontal="center" if spalte > start else "left",
                                vertical="center", wrap_text=True)


def _link(zelle, ziel: str) -> None:
    zelle.hyperlink = Hyperlink(ref=zelle.coordinate, location=ziel, display=str(zelle.value))
    zelle.font = FONT_LINK


def _druck(ws, titelzeilen: str = None) -> None:
    """Querformat, eine Seite breit, Blattname und Seitenzahl in der Fußzeile."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Seite &P von &N"
    if titelzeilen:
        ws.print_title_rows = titelzeilen


# --- Blatt Überblick ---

def _blatt_ueberblick(ex, wb, quelle: str, blaetter: list) -> None:
    ws = ex.active
    ws.title = UEBERBLICK
    basis = _wert(wb, "par_Basisjahr")
    jahre = _bereich(wb, "liq_Jahr")
    zeile = _titel(ws, "Prognose Immobilien – Überblick",
                   f"Basisjahr {basis} (Ist), Prognose {jahre[0]}–{jahre[-1]}. Erstellt am "
                   f"{datetime.date.today():%d.%m.%Y}" + (f" aus {quelle}" if quelle else "")
                   + ". Alle Werte beruhen auf den Annahmen im Blatt Annahmen; die Steuer ist "
                   "mit dem Grenzsteuersatz gerechnet.")

    _abschnitt(ws, zeile, "Ergebnis", 4)
    zeile += 1
    for name, font in (("start_Empfehlung", Font(bold=True, size=12, color="305496")),
                       ("start_Belastbarkeit", None)):
        c = ws.cell(row=zeile, column=1, value=_wert(wb, name))
        if font:
            c.font = font
        zeile += 1
    zeile += 1
    _kopfzeile(ws, zeile, [f"Option (Endvermögen Ende {jahre[-1]})",
                           "Endvermögen nach latenter Steuer", "Differenz zu Halten", "Rang"])
    zeile += 1
    werte = [_zahl(_wert(wb, name)) for _, name in OPTIONEN]
    for (text, _), wert in zip(OPTIONEN, werte):
        rang = sum(w > wert + SCHWELLE for w in werte) + 1
        ws.cell(row=zeile, column=1, value=text).font = FONT_FETT if rang == 1 else Font()
        for spalte, inhalt, fmt in ((2, wert, FMT_BETRAG), (3, wert - werte[0], FMT_DIFF),
                                    (4, rang, "0")):
            ws.cell(row=zeile, column=spalte, value=inhalt).number_format = fmt
        zeile += 1
    zeile += 1
    kum = _bereich(wb, "liq_Kum")
    tiefst = min(range(len(kum)), key=lambda i: _zahl(kum[i]))
    for text, wert, fmt in [
        ("Wert der § 6b-Kette (Plan gegen sofort versteuern und kaufen)",
         _wert(wb, "vg_DiffC"), FMT_DIFF),
        ("Steuer gesamt im Plan über alle Jahre", sum(map(_zahl, _bereich(wb, "liq_Steuer"))),
         FMT_BETRAG),
        (f"tiefster Liquiditätsstand im Plan (im Jahr {jahre[tiefst]}; negativ = "
         "Finanzierungsbedarf)", kum[tiefst], FMT_BETRAG),
    ]:
        ws.cell(row=zeile, column=1, value=text)
        ws.cell(row=zeile, column=2, value=wert).number_format = fmt
        zeile += 1

    zeile += 1
    _abschnitt(ws, zeile, "Datenlage", 4)
    zeile += 1
    for text, name in (("Objekte im Modell", "ueb_Objekte"),
                       ("davon mit Annahmen (Blatt Annahmen, blau)", "ueb_MitAnnahmen"),
                       ("geplante Verkäufe", "ueb_Verkaeufe"),
                       ("Neuobjekte", "ueb_Neuobjekte"),
                       ("Plausibilitätsprüfung", "pr_Gesamt")):
        ws.cell(row=zeile, column=1, value=text)
        ws.cell(row=zeile, column=2, value=_wert(wb, name)).alignment = Alignment(
            horizontal="right")
        zeile += 1

    zeile += 1
    _abschnitt(ws, zeile, "Inhalt", 4)
    zeile += 1
    for blatt, text in INHALT:
        if blatt in blaetter:
            _link(ws.cell(row=zeile, column=1, value=blatt), f"'{blatt}'!A1")
            ws.cell(row=zeile, column=2, value=text)
            zeile += 1

    zeile += 1
    _abschnitt(ws, zeile, "Entwicklung je Jahr (Ende des Jahres)", 7)
    zeile += 1
    ws.cell(row=zeile, column=2, value="Immobilien (Verkehrswert)").font = FONT_FETT
    ws.cell(row=zeile, column=5, value="Gesamtvermögen nach latenter Steuer").font = FONT_FETT
    kopf = zeile + 1
    _kopfzeile(ws, kopf, ["Jahr", "Halten", "Plan", "Differenz Plan − Halten",
                          "Halten", "Plan", "Differenz Plan − Halten"])
    spalten = [_bereich(wb, n) for n in ("ueb_Jahr", "ueb_Baseline", "ueb_Plan",
                                         "ueb_Differenz", "ueb_VermBaseline", "ueb_VermPlan",
                                         "ueb_VermDifferenz")]
    erste = kopf + 1
    for i, werte_jahr in enumerate(zip(*spalten)):
        for spalte, wert in enumerate(werte_jahr, start=1):
            c = ws.cell(row=erste + i, column=spalte, value=wert)
            c.number_format = "0" if spalte == 1 else (
                FMT_DIFF if spalte in (4, 7) else FMT_BETRAG)
            if spalte == 1:
                c.alignment = Alignment(horizontal="left")
    letzte = erste + len(spalten[0]) - 1
    # Diagramme rechts neben Ergebnis und Datenlage, die nur die Spalten A bis D nutzen
    for i, (titel, spalte) in enumerate((("Gesamtvermögen nach latenter Steuer", 5),
                                         ("Verkehrswert am Jahresende", 2))):
        chart = _diagramm(ws, titel, spalte, kopf, erste, letzte)
        chart.height, chart.width = 7, 17
        ws.add_chart(chart, f"F{4 + i * 14}")

    ws.column_dimensions["A"].width = 62
    for spalte in "BCDEFG":
        ws.column_dimensions[spalte].width = 17
    _druck(ws)


# --- Blatt Jahresrechnung ---

@dataclass(frozen=True)
class Posten:
    text: str
    quelle: str = None      # "liq" (Blatt Liquidität) oder "aus" (Blatt Auswertung)
    feld: str = None        # Name ohne Präfix, z. B. "Einnahmen"
    vorzeichen: int = 1     # −1: in der Quelle positiv, hier abgezogen
    art: str = "fluss"      # fluss: Summe über die Jahre, summe: dazu fett, bestand: Stand


JAHRESRECHNUNG_POSTEN = [
    Posten("Ergebnis und Steuer"),
    Posten("Mieten und weitere Einnahmen", "liq", "Einnahmen"),
    Posten("./. Erhaltung und weitere Ausgaben", "liq", "Ausgaben", -1),
    Posten("./. Abschreibungen (AfA)", "liq", "AfA", -1),
    Posten("= laufendes Ergebnis", "liq", "Ergebnis", art="summe"),
    Posten("+ steuerwirksam aus Verkauf und § 6b-Rücklage", "liq", "Verkauf"),
    Posten("+ Zinsertrag Alternativanlage", "liq", "Zins"),
    Posten("./. Zinsen Darlehen", "liq", "KreditZins", -1),
    Posten("= Ergebnis vor Verlustvortrag", "liq", "ZvE", art="summe"),
    Posten("./. Verlustvortrag genutzt", "liq", "VortragGenutzt", -1),
    Posten("= Bemessungsgrundlage", "liq", "Bemessung", art="summe"),
    Posten("Steuer", "liq", "Steuer", art="summe"),
    Posten("Verlustvortrag am Jahresende", "liq", "Vortrag", art="bestand"),
    Posten("Geldfluss"),
    Posten("laufendes Ergebnis", "liq", "Ergebnis"),
    Posten("+ Abschreibungen (nicht zahlungswirksam)", "liq", "AfA"),
    Posten("+ Zinsertrag Alternativanlage", "liq", "Zins"),
    Posten("./. Zinsen Darlehen", "liq", "KreditZins", -1),
    Posten("./. Steuer", "liq", "Steuer", -1),
    Posten("+ Verkaufserlöse netto", "liq", "Verkaufserloes"),
    Posten("./. Kauf Neuobjekte inkl. Nebenkosten", "liq", "Kauf", -1),
    Posten("+ Kreditauszahlung", "liq", "Kredit"),
    Posten("./. Tilgung Darlehen (mit Ablösung bei Verkauf)", "liq", "Tilgung", -1),
    Posten("= freier Mittelzufluss", "liq", "Zufluss", art="summe"),
    Posten("Liquidität am Jahresende (Alternativanlage)", "liq", "Kum", art="bestand"),
    Posten("Vermögen am Jahresende"),
    Posten("Verkehrswert Immobilien", "aus", "Verkehrswert", art="bestand"),
    Posten("+ Liquidität", "aus", "Liquiditaet", art="bestand"),
    Posten("./. Restschuld Darlehen", "aus", "Restschuld", -1, art="bestand"),
    Posten("= Gesamtvermögen vor latenter Steuer", "aus", "Vermoegen", art="bestand"),
    Posten("./. latente Steuer auf stille Reserven und Rücklage", "aus", "LatenteSteuer", -1,
           art="bestand"),
    Posten("= Gesamtvermögen nach latenter Steuer", "aus", "VermoegenNetto", art="bestand"),
    Posten("nachrichtlich: Buchwert Immobilien", "aus", "Buchwert", art="bestand"),
    Posten("nachrichtlich: stille Reserven", "aus", "StilleReserven", art="bestand"),
    Posten("nachrichtlich: § 6b-Rücklage", "aus", "Ruecklage", art="bestand"),
]
DIFFERENZ_POSTEN = [
    Posten("laufendes Ergebnis", "liq", "Ergebnis"),
    Posten("Steuer", "liq", "Steuer"),
    Posten("freier Mittelzufluss", "liq", "Zufluss"),
    Posten("Liquidität am Jahresende", "liq", "Kum", art="bestand"),
    Posten("Verkehrswert Immobilien", "aus", "Verkehrswert", art="bestand"),
    Posten("Restschuld Darlehen", "aus", "Restschuld", art="bestand"),
    Posten("Gesamtvermögen nach latenter Steuer", "aus", "VermoegenNetto", art="bestand"),
]
# Präfixe der Namen je Sicht: Plan = Szenario A, Halten = Baseline
SICHTEN = {"Plan": {"liq": "liq", "aus": "aus"}, "Halten": {"liq": "lqb", "aus": "asb"}}


def _reihe(wb, sicht: str, p: Posten) -> list:
    return [p.vorzeichen * _zahl(w) for w in _bereich(wb, f"{SICHTEN[sicht][p.quelle]}_{p.feld}")]


def _jahreszeilen(ws, zeile: int, posten: list, reihen, n: int, fmt: str = FMT_BETRAG) -> int:
    """Je Posten eine Zeile: Text, Jahreswerte, Summe (nur Flussgrößen)."""
    for p in posten:
        if p.quelle is None:
            ws.cell(row=zeile, column=1, value=p.text).font = FONT_FETT
            zeile += 1
            continue
        werte = reihen(p)
        ws.cell(row=zeile, column=1, value=p.text)
        for i, wert in enumerate(werte, start=2):
            ws.cell(row=zeile, column=i, value=wert).number_format = fmt
        if p.art != "bestand":
            ws.cell(row=zeile, column=n + 2, value=sum(werte)).number_format = fmt
        if p.text.startswith("="):
            for spalte in range(1, n + 3):
                c = ws.cell(row=zeile, column=spalte)
                c.font, c.border = FONT_FETT, RAHMEN_SUMME
        zeile += 1
    return zeile


def _blatt_jahresrechnung(ex, wb) -> None:
    ws = ex.create_sheet(JAHRESRECHNUNG)
    jahre = _bereich(wb, "liq_Jahr")
    n = len(jahre)
    zeile = _titel(ws, "Jahresrechnung: vom Ergebnis zum Vermögen",
                   "Plan = mit den geplanten Verkäufen und Neuobjekten (§ 6b wie erfasst), "
                   "Halten = alle Objekte behalten. Beträge in €, „./.“ wird abgezogen; jede "
                   "Spalte rechnet von oben nach unten. Summe = über alle Jahre.")
    _kopfzeile(ws, zeile, ["Position"] + jahre + ["Summe"])
    ws.freeze_panes = ws.cell(row=zeile + 1, column=2)
    kopf = zeile
    zeile += 1
    plan_gleich_halten = all(
        all(abs(a - b) <= SCHWELLE for a, b in zip(_reihe(wb, "Plan", p), _reihe(wb, "Halten", p)))
        for p in JAHRESRECHNUNG_POSTEN if p.quelle)
    for sicht, titel in (("Plan", "Plan: mit Verkäufen und Neuobjekten"),
                         ("Halten", "Halten: alle Objekte behalten")):
        if sicht == "Halten" and plan_gleich_halten:
            break
        zeile += 1
        _abschnitt(ws, zeile, titel, n + 2)
        zeile = _jahreszeilen(ws, zeile + 1, JAHRESRECHNUNG_POSTEN,
                              lambda p, s=sicht: _reihe(wb, s, p), n)
    zeile += 1
    if plan_gleich_halten:
        ws.cell(row=zeile, column=1, value="Ohne Verkäufe und Neuobjekte gleicht der Plan dem "
                                           "Halten.").font = FONT_HINWEIS
    else:
        _abschnitt(ws, zeile, "Differenz Plan − Halten (+ Plan besser)", n + 2)
        zeile = _jahreszeilen(
            ws, zeile + 1, DIFFERENZ_POSTEN,
            lambda p: [a - b for a, b in zip(_reihe(wb, "Plan", p), _reihe(wb, "Halten", p))],
            n, FMT_DIFF)
    ws.column_dimensions["A"].width = BREITE_TEXT
    for spalte in range(2, n + 3):
        ws.column_dimensions[get_column_letter(spalte)].width = BREITE_ZAHL
    _druck(ws, f"{kopf}:{kopf}")


# --- Blatt Ergebnisbrücke ---

def _blatt_bruecke(ex, wb) -> bool:
    """Werte des Blatts Ergebnisbrücke ohne leere Pakete und Vorgänge, ohne Paketspalte und
    Kontrollzeile; ohne
    Verkauf und Kauf kein Blatt (False)."""
    if not any(_ungleich_null(_bereich(wb, f"brk_{name}"))
               for name in ("Veraeusserung", "Laufend", "Neuobjekte")):
        return False
    quelle = wb[BRUECKE]
    blatt, ref = next(iter(wb.defined_names["brk_Halten"].destinations))
    erste = wb[blatt][ref.replace("$", "").split(":")[0]]
    kopf, spalte = erste.row - 1, erste.column
    jahre = [c.value for c in quelle[kopf][spalte - 1:]
             if isinstance(c.value, (int, float))]
    n = len(jahre)
    kontrolle = next(iter(wb.defined_names["brk_Kontrolle"].destinations))[1]
    kontrollzeile = int(kontrolle.split("$")[2].split(":")[0])
    ws = ex.create_sheet(BRUECKE)
    zeile = _titel(ws, "Ergebnisbrücke: Wirkung der Entscheidungen",
                   "Vom Ergebnis vor Steuern bei Halten je Verkauf und Kauf zum Ergebnis vor "
                   "Steuern im Plan. Verkauf: Gewinn und § 6b-Rücklage minus das laufende "
                   "Ergebnis, das mit dem Objekt entfällt. Kauf: Miete − Kosten − AfA − "
                   "Kreditzinsen. Summe = über alle Jahre.")
    _kopfzeile(ws, zeile, ["Position"] + jahre + ["Summe"])
    ws.freeze_panes = ws.cell(row=zeile + 1, column=2)
    kopfzeile = zeile
    zeile += 1
    for q in range(kopf + 1, quelle.max_row + 1):
        text = quelle.cell(row=q, column=1).value
        if text in (None, "") or q == kontrollzeile:
            continue
        werte = [_zahl(quelle.cell(row=q, column=spalte + i).value) for i in range(n)]
        einzug = quelle.cell(row=q, column=1).alignment.indent
        if quelle.cell(row=q, column=2).value is None:   # Abschnitt
            zeile += 1
            _abschnitt(ws, zeile, text, n + 2)
            zeile += 1
            continue
        # leere Pakete und Vorgänge entfallen, feste Summenzeilen bleiben
        if not _ungleich_null(werte) and not text.startswith(
                ("Ergebnis", "Summe", "Zinseffekt", "Wirkung")):
            continue
        fett = bool(quelle.cell(row=q, column=1).font.bold)
        c = ws.cell(row=zeile, column=1, value=text)
        c.alignment = Alignment(indent=einzug)
        for i, wert in enumerate(werte + [sum(werte)], start=2):
            ws.cell(row=zeile, column=i, value=wert).number_format = FMT_BETRAG
        if fett:
            for i in range(1, n + 3):
                c = ws.cell(row=zeile, column=i)
                c.font, c.border = FONT_FETT, RAHMEN_SUMME
        zeile += 1
    ws.column_dimensions["A"].width = BREITE_TEXT
    for i in range(2, n + 3):
        ws.column_dimensions[get_column_letter(i)].width = BREITE_ZAHL
    _druck(ws, f"{kopfzeile}:{kopfzeile}")
    return True


# --- BWA-Blätter ---

@dataclass
class BwaZeile:
    art: str          # bwa (Zeile mit BWA-Nr.), titel (Abschnitt), daten (Herleitung u. Ä.)
    nr: int
    text: str
    ziel: object      # Spalte D: BWA-Nr. des Postens bzw. Anlagennummer
    werte: list


def ist_bwa_blatt(ws) -> bool:
    return ws["B4"].value == "Nr." and ws.cell(row=ERSTE_ZEILE - 1, column=3).value \
        == "Jahr (für Formeln)"


def _bwa_spalten(wb) -> list:
    """(Spalte im BWA-Blatt, Kopf im Export): Ist-Vorjahre nur, wenn eingelesen."""
    basis = _wert(wb, "par_Basisjahr")
    summe = _bwa_lesen(wb[SUMMENBLATT], [(SPALTE_VORJAHRE, ""), (SPALTE_VORJAHRE + 1, "")])
    vorjahre = [(SPALTE_VORJAHRE + i, f"Ist {basis - 2 + i}") for i in range(2)
                if _ungleich_null(z.werte[i] for z in summe if z.art == "bwa")]
    plan = [(SPALTE_PLAN + i, jahr) for i, jahr in enumerate(_bereich(wb, "liq_Jahr"))]
    return vorjahre + [(SPALTE_JAHR, f"Ist {basis}")] + plan


def _bwa_lesen(ws, spalten: list) -> list:
    """Zeilen eines BWA-Blatts ab der ersten BWA-Zeile: Werte der gewählten Spalten."""
    zeilen = []
    for zeile in ws.iter_rows(min_row=ERSTE_ZEILE, values_only=True):
        zeile = zeile + (None,) * (max(s for s, _ in spalten) - len(zeile))
        nr, text, ziel = zeile[1], zeile[2], zeile[3]
        if not isinstance(text, str) or not text.strip():
            continue
        werte = [zeile[s - 1] for s, _ in spalten]
        if isinstance(nr, int) and not text.endswith(":"):
            art = "bwa"
        elif isinstance(nr, int) or not any(isinstance(w, (int, float)) for w in werte):
            art = "titel"
        else:
            art = "daten"
        zeilen.append(BwaZeile(art, nr if isinstance(nr, int) else None, text, ziel, werte))
    return zeilen


def _bwa_filtern(zeilen: list) -> list:
    """Zeilen ohne Wert weglassen; Summenzeilen bleiben, Abschnitte nur mit Inhalt."""
    ergebnis, abschnitt = [], []

    def abschliessen():
        if any(z.art != "titel" for z in abschnitt):
            ergebnis.extend(abschnitt)
        abschnitt.clear()

    for z in zeilen:
        if z.art == "bwa":
            abschliessen()
            if z.nr in SUMMENZEILEN or _ungleich_null(z.werte):
                ergebnis.append(z)
        elif z.art == "titel":
            abschliessen()
            abschnitt.append(z)
        elif _ungleich_null(z.werte):
            abschnitt.append(z)
    abschliessen()
    return ergebnis


def _bwa_block(ws, zeile: int, zeilen: list, spalten: list, fmt: str = FMT_BWA) -> int:
    """Kopf und Zeilen einer BWA ab zeile; liefert die nächste freie Zeile."""
    _kopfzeile(ws, zeile, ["Bezeichnung", "Nr. / →"] + [k for _, k in spalten])
    zeile += 1
    for z in zeilen:
        ws.cell(row=zeile, column=1, value=z.text)
        ziel = z.nr if z.art == "bwa" else z.ziel
        if ziel not in (None, "", "→ Nr."):
            ws.cell(row=zeile, column=2, value=ziel).alignment = Alignment(horizontal="center")
        if z.art == "titel":
            ws.cell(row=zeile, column=1).font = FONT_FETT
        else:
            for i, wert in enumerate(z.werte, start=3):
                if isinstance(wert, (int, float)):
                    ws.cell(row=zeile, column=i, value=wert).number_format = fmt
        if z.nr in SUMMENZEILEN:
            for spalte in range(1, len(spalten) + 3):
                c = ws.cell(row=zeile, column=spalte)
                c.font, c.border = FONT_FETT, RAHMEN_SUMME
        zeile += 1
    return zeile


def _bwa_breiten(ws, spalten: list) -> None:
    ws.column_dimensions["A"].width = BREITE_TEXT
    ws.column_dimensions["B"].width = 9
    for i in range(len(spalten)):
        ws.column_dimensions[get_column_letter(i + 3)].width = BREITE_ZAHL


def _blatt_bwa_gesamt(ex, wb, spalten: list) -> None:
    ws = ex.create_sheet(BWA_GESAMT)
    vergleich = str(_wert(wb, "bwa_VglHalten")).strip().lower() == "ja"
    zeile = _titel(ws, "BWA gesamt: alle Objekte",
                   "Ist-Spalten aus der Buchhaltung, Planspalten aus der Prognose (mit Verkäufen "
                   "und Neuobjekten), Steuer und Zinsertrag der Alternativanlage nur hier. "
                   "Unter der BWA die Herleitung der Sonderposten mit der BWA-Nr., in die sie "
                   "eingehen. Zeilen ohne Wert sind ausgelassen.")
    plan = _bwa_lesen(wb[SUMMENBLATT], spalten)
    _abschnitt(ws, zeile, "Plan: mit Verkäufen und Neuobjekten", len(spalten) + 2)
    zeile = _bwa_block(ws, zeile + 1, _bwa_filtern(plan), spalten) + 1
    # Halten selbst steht im Blatt BWA halten
    if vergleich:
        _abschnitt(ws, zeile, f"Differenz Plan − Halten (+ Plan höher; Halten im Blatt "
                              f"{BWA_HALTEN})", len(spalten) + 2)
        halten = {z.nr: z.werte for z in _bwa_lesen(wb[HALTEBLATT], spalten) if z.art == "bwa"}
        plan_ab = _plan_ab(spalten)
        diff = [BwaZeile("bwa", z.nr, z.text, None,
                         [None] * plan_ab + [_zahl(a) - _zahl(b) for a, b in
                                             zip(z.werte[plan_ab:], halten[z.nr][plan_ab:])])
                for z in plan if z.art == "bwa" and z.nr in halten]
        _bwa_block(ws, zeile + 1, _bwa_filtern(diff), spalten, FMT_DIFF)
    _bwa_breiten(ws, spalten)
    ws.freeze_panes = "C5"
    _druck(ws)


def _plan_ab(spalten: list) -> int:
    """Index der ersten Planspalte in spalten."""
    return next(i for i, (s, _) in enumerate(spalten) if s >= SPALTE_PLAN)


def _blatt_bwa_halten(ex, wb, spalten: list) -> bool:
    """Summen-BWA bei Halten (Status quo). False und kein Blatt, wenn sie in den Planjahren
    der BWA gesamt gleicht, also ohne Verkauf und Kauf."""
    halten = _bwa_lesen(wb[HALTEBLATT], spalten)
    plan = {z.nr: z.werte for z in _bwa_lesen(wb[SUMMENBLATT], spalten) if z.art == "bwa"}
    ab = _plan_ab(spalten)
    if all(abs(_zahl(a) - _zahl(b)) <= SCHWELLE
           for z in halten if z.art == "bwa"
           for a, b in zip(z.werte[ab:], plan.get(z.nr, [None] * len(spalten))[ab:])):
        return False
    ws = ex.create_sheet(BWA_HALTEN)
    zeile = _titel(ws, "BWA halten: Status quo, alle Objekte behalten",
                   "Wie die BWA gesamt, aber ohne Verkäufe und Neuobjekte: alle Bestandsobjekte "
                   "bleiben über das ganze Raster, ihre Darlehen laufen weiter. Ist-Spalten wie "
                   "in der BWA gesamt, Planspalten aus der Baseline (Steuer mit Verlustvortrag). "
                   "Zeilen ohne Wert sind ausgelassen.")
    _abschnitt(ws, zeile, "Halten: alle Objekte behalten", len(spalten) + 2)
    _bwa_block(ws, zeile + 1, _bwa_filtern(halten), spalten)
    _bwa_breiten(ws, spalten)
    ws.freeze_panes = "C5"
    _druck(ws)
    return True


def _blatt_bwa_objekte(ex, wb, spalten: list) -> None:
    ws = ex.create_sheet(BWA_OBJEKTE)
    blaetter = [b for b in wb.worksheets
                if b.title not in (SUMMENBLATT, HALTEBLATT) and ist_bwa_blatt(b)]
    zeile = _titel(ws, "BWA je Objekt",
                   "Je Kostenstelle die BWA: Ist-Spalten aus der Buchhaltung, Planspalten aus "
                   "der Prognose; nach einem Verkauf leer, Neuobjekte ab dem Jahr nach dem Kauf. "
                   "Steuer und Zinsertrag der Alternativanlage stehen nur im Blatt BWA gesamt.")
    _abschnitt(ws, zeile, "Inhalt", 2)
    inhalt = zeile + 1
    zeile = inhalt + len(blaetter) + 1
    for i, quelle in enumerate(blaetter):
        kst, name = quelle["B2"].value, quelle["C2"].value
        titel = f"{kst} – {name}" if name else str(kst)
        _link(ws.cell(row=inhalt + i, column=1, value=titel), f"'{BWA_OBJEKTE}'!A{zeile}")
        _abschnitt(ws, zeile, titel, len(spalten) + 2)
        zeile = _bwa_block(ws, zeile + 1, _bwa_filtern(_bwa_lesen(quelle, spalten)),
                           spalten) + 1
    if not blaetter:
        ws.cell(row=inhalt, column=1, value="keine Kostenstellen").font = FONT_HINWEIS
    _bwa_breiten(ws, spalten)
    ws.freeze_panes = "C4"
    _druck(ws)


# --- Blatt Verkauf und Kauf ---

VERKAUF_SPALTEN = [("ObjektID", "vk_ID", None), ("weitere Kostenstellen", "vk_Weitere", None),
                   ("Verkauf aktiv", "vk_Aktiv", None),
                   ("Verkaufsjahr", "vk_Jahr", "0"), ("Verkaufspreis", "vk_Preis", FMT_BETRAG),
                   ("Verkaufskosten", "vk_Kosten", FMT_BETRAG),
                   ("Nettoerlös", "vk_Nettoerloes", FMT_BETRAG),
                   ("Buchwert Gebäude Ende Verkaufsjahr", "vk_BuchwertGeb", FMT_BETRAG),
                   ("AK G+B", "vk_AKGuB", FMT_BETRAG),
                   ("Veräußerungsgewinn", "vk_Gewinn", FMT_BETRAG),
                   ("§ 6b nutzen", "vk_6b", None), ("reinvestieren", "vk_Reinvest", None),
                   ("Status", "vk_Status", None)]
NEU_SPALTEN = [("NeuID", "ne_ID", None), ("Kauf aktiv", "ne_Aktiv", None),
               ("Name", "ne_Name", None), ("Kaufjahr", "ne_Kaufjahr", "0"),
               ("Kaufpreis", "ne_Kaufpreis", FMT_BETRAG),
               ("Kaufpreis + Nebenkosten", "ne_AKGesamt", FMT_BETRAG),
               ("Übertrag § 6b", "ne_UeGesamt", FMT_BETRAG),
               ("AfA-Basis Gebäude", "ne_AfABasis", FMT_BETRAG),
               ("Kredit", "ne_Kredit", FMT_BETRAG), ("Eigenmittel", "ne_Eigen", FMT_BETRAG),
               ("Status", "ne_Status", None)]
SB_AUSLASSEN = ("rechnet mit (1 = ja)",)
GRAU = "E7E6E6"   # FILL_BERECHNET der Mappe


def _tabelle(ws, zeile: int, wb, spalten: list) -> tuple:
    """Eingabezeilen mit ID als Tabelle; ausgeschaltete grau. Liefert (nächste Zeile, Anzahl)."""
    reihen = [_bereich(wb, name) for _, name, _ in spalten]
    _kopfzeile(ws, zeile, [k for k, _, _ in spalten])
    zeile += 1
    anzahl = 0
    for werte in zip(*reihen):
        if werte[0] in (None, ""):
            continue
        aus = str(werte[1]).strip().lower() == "nein"
        for spalte, ((_, _, fmt), wert) in enumerate(zip(spalten, werte), start=1):
            if spalte == 2 and wert in (None, ""):
                wert = "ja"
            c = ws.cell(row=zeile, column=spalte, value=wert)
            if fmt:
                c.number_format = fmt
            if aus:
                c.font = Font(color="808080", strike=True)
        zeile += 1
        anzahl += 1
    return zeile, anzahl


def _sonderbereich(ws, zeile: int, quelle) -> int:
    """Blöcke je Vorgang und Detailsicht Kauf aus dem Blatt Verkauf und Kauf, nur mit Inhalt."""
    zeilen = list(quelle.iter_rows(min_row=1, max_col=5))
    bloecke, aktuell = [], None
    for z in zeilen:
        text = z[0].value
        if isinstance(text, str) and (text.startswith("Vorgang ")
                                      or text.startswith("Detailsicht Kauf")):
            aktuell = []
            bloecke.append(aktuell)
        elif isinstance(text, str) and text.startswith("Weitere Vorgänge"):
            aktuell = None
        if aktuell is not None:
            aktuell.append(z)
    for block in bloecke:
        titel = block[0][0].value
        if titel.startswith("Vorgang "):
            mit = next((z[1].value for z in block if z[0].value == SB_AUSLASSEN[0]), 0)
            if not _zahl(mit):
                continue
            spalten = range(5)
        else:   # Detailsicht Kauf: nur Spalten der erfassten Neuobjekte
            ids = next((z for z in block if z[0].value == "NeuID (Kostenstelle)"), None)
            spalten = [0] + [i for i in range(1, 5) if ids and ids[i].value not in (None, "")]
            if len(spalten) == 1:
                continue
        for z in block:
            if z[0].value in SB_AUSLASSEN:
                continue
            for ziel, i in enumerate(spalten, start=1):
                q = z[i]
                c = ws.cell(row=zeile, column=ziel, value=q.value)
                if q.has_style:
                    c.font, c.alignment = copy(q.font), copy(q.alignment)
                    if not str(q.fill.fgColor.rgb).endswith(GRAU):   # grau = Formel
                        c.fill = copy(q.fill)
                    c.number_format = (FMT_BETRAG if q.number_format == FMT_EURO
                                       else q.number_format)
            zeile += 1
        zeile += 1
    return zeile


def _blatt_verkauf_kauf(ex, wb) -> bool:
    """Verkäufe, Neuobjekte und Sonderbereich; ohne beides kein Blatt (False)."""
    ws = ex.create_sheet(VERKAUF_KAUF)
    zeile = _titel(ws, "Verkauf und Kauf",
                   "Geplante Verkäufe und Neuobjekte; darunter je Verkauf das Ergebnis "
                   "(Erlös, Reinvestition, Kapitalanlage, Vergleich mit Halten) und die "
                   "Herleitung von Gewinn und § 6b-Rücklage, dann je Neuobjekt Kauf und "
                   "Finanzierung. Ausgeschaltete Vorgänge sind durchgestrichen und rechnen "
                   "nicht mit.")
    _abschnitt(ws, zeile, "Verkäufe", len(VERKAUF_SPALTEN))
    zeile, verkaeufe = _tabelle(ws, zeile + 1, wb, VERKAUF_SPALTEN)
    zeile += 1
    _abschnitt(ws, zeile, "Neuobjekte", len(VERKAUF_SPALTEN))
    zeile, neu = _tabelle(ws, zeile + 1, wb, NEU_SPALTEN)
    if not verkaeufe and not neu:
        ex.remove(ws)
        return False
    _sonderbereich(ws, zeile + 1, wb[SONDERBEREICH])
    ws.column_dimensions["A"].width = 66
    for spalte in range(2, len(VERKAUF_SPALTEN) + 1):
        ws.column_dimensions[get_column_letter(spalte)].width = 16
    _druck(ws)
    return True


# --- Blatt Annahmen ---

OBJEKT_SPALTEN = ("objekt_id", "name", "miete", "erhaltung", "weitere_einnahmen",
                  "weitere_ausgaben", "afa_bwa", "zinsen", "verkehrswert", "vk_quote_gebaeude",
                  "kaufjahr", "baujahr", "ak_gebaeude", "ak_gub", "restbuchwert", "afa_satz",
                  "afa_jahr", "restschuld", "zinssatz", "rate")
ANNAHME, BUCHHALTUNG = "Annahme", "Buchhaltung"
HERKUNFT_FILL = {ANNAHME: FILL_ANNAHME, BUCHHALTUNG: FILL_EINGELESEN}


def _spalte(buchstabe: str) -> int:
    return column_index_from_string(buchstabe) - 1


def herkunft(formel_zeile: tuple, wert_zeile: tuple) -> dict:
    """Herkunft je Objektfeld wie die Farblogik des Blatts Objekte: Annahme (blau, Formel der
    Auffülllogik), Buchhaltung (grün, Kostenstellenblatt, eingelesen oder Blatt Anlagen) oder
    keine Angabe (eingetippt)."""
    def zahl(k):
        return _zahl(wert_zeile[_spalte(hspalte(k))])

    ergebnis = {}
    restschuld = 0
    for i, f in enumerate(OBJEKT_FELDER):
        inhalt = formel_zeile[i] if i < len(formel_zeile) else None
        eingelesen = (wert_zeile[_spalte(ispalte(f.key))]
                      if f.key in OBJEKT_EINGELESEN else None)
        if f.key == "restschuld":
            restschuld = _zahl(wert_zeile[i])
        if isinstance(inhalt, str) and inhalt.startswith("="):
            if f.key in KST_VERKNUEPFT and eingelesen not in (None, ""):
                ergebnis[f.key] = BUCHHALTUNG
            elif deckt_anlagen(f.key, zahl):
                ergebnis[f.key] = BUCHHALTUNG
            elif f.key in DARLEHEN_ANNAHMEN:
                if restschuld > 0:
                    ergebnis[f.key] = ANNAHME
            elif f.annahme:
                ergebnis[f.key] = ANNAHME
        elif eingelesen not in (None, "") and inhalt == eingelesen:
            ergebnis[f.key] = BUCHHALTUNG
    return ergebnis


def _objektzeilen(ws, anzahl: int) -> list:
    return [tuple(z) for z in ws.iter_rows(min_row=2, max_row=anzahl + 1, values_only=True)]


def _blatt_annahmen(ex, wb, formeln_wb) -> None:
    ws = ex.create_sheet(ANNAHMEN)
    zeile = _titel(ws, "Annahmen und Datenlage",
                   "Die zentralen Annahmen gelten für alle Objekte, soweit keine Daten aus der "
                   "Buchhaltung oder eigene Werte vorliegen. Je Objekt zeigt die Farbe, woher "
                   "ein Wert stammt.")
    _abschnitt(ws, zeile, "Zentrale Annahmen", 2)
    zeile += 1
    for p in [q for q in PARAMETER if q.name in ("par_Basisjahr", "par_Prognosejahre")] \
            + zentrale_annahmen():
        ws.cell(row=zeile, column=1, value=p.bezeichnung)
        c = ws.cell(row=zeile, column=2, value=_wert(wb, p.name))
        c.number_format = FMT_BETRAG if p.format == FMT_EURO else p.format
        c.alignment = Alignment(horizontal="right")
        zeile += 1

    zeile += 1
    felder = [f for f in OBJEKT_FELDER if f.key in OBJEKT_SPALTEN]
    _abschnitt(ws, zeile, "Daten je Objekt (Basisjahr)", len(felder) + 1)
    zeile += 1
    for fill, text in ((FILL_ANNAHME, "Annahme: fehlende Angabe, nach den zentralen Annahmen "
                                      "ergänzt"),
                       (FILL_EINGELESEN, "aus der Buchhaltung (BWA) oder dem "
                                         "Anlagenverzeichnis"),
                       (None, "ohne Farbe: eigens erfasster Wert")):
        c = ws.cell(row=zeile, column=1, value=text)
        if fill:
            c.fill = fill
        zeile += 1
    _kopfzeile(ws, zeile, [f.ueberschrift.rstrip(" *") for f in felder] + ["Status"])
    ws.row_dimensions[zeile].height = 45
    zeile += 1
    anzahl = len(_bereich(wb, "obj_ID"))
    werte_zeilen = _objektzeilen(wb["Objekte"], anzahl)
    formel_zeilen = (_objektzeilen(formeln_wb["Objekte"], anzahl) if formeln_wb is not None
                     else [()] * anzahl)
    status = _bereich(wb, "obj_Status")
    for wert_zeile, formel_zeile, st in zip(werte_zeilen, formel_zeilen, status):
        if wert_zeile[0] in (None, ""):
            continue
        quelle = herkunft(formel_zeile, wert_zeile) if formel_zeile else {}
        for spalte, f in enumerate(felder, start=1):
            wert = wert_zeile[OBJEKT_FELDER.index(f)]
            c = ws.cell(row=zeile, column=spalte, value=wert)
            c.number_format = FMT_BETRAG if f.format == FMT_EURO else f.format
            if f.key in quelle:
                c.fill = HERKUNFT_FILL[quelle[f.key]]
        ws.cell(row=zeile, column=len(felder) + 1, value=st)
        zeile += 1

    zeile += 1
    _abschnitt(ws, zeile, f"Plausibilitätsprüfung: {_wert(wb, 'pr_Gesamt')}", 3)
    zeile += 1
    auffaellig = [(text, art, n) for text, art, n in zip(_bereich(wb, "pr_Bezeichnung"),
                                                         _bereich(wb, "pr_Art"),
                                                         _bereich(wb, "pr_Anzahl"))
                  if _zahl(n) > 0]
    if auffaellig:
        _kopfzeile(ws, zeile, ["Prüfung", "Art", "betroffen"])
        zeile += 1
        for text, art, n in auffaellig:
            ws.cell(row=zeile, column=1, value=text).alignment = Alignment(wrap_text=True)
            ws.cell(row=zeile, column=2, value=art)
            ws.cell(row=zeile, column=3, value=n)
            zeile += 1
    else:
        ws.cell(row=zeile, column=1, value="keine Auffälligkeiten").font = FONT_HINWEIS
    ws.column_dimensions["A"].width = BREITE_TEXT
    for spalte in range(2, len(felder) + 2):
        ws.column_dimensions[get_column_letter(spalte)].width = 14
    _druck(ws)


# --- Mappe ---

def erstelle_export(wb, formeln_wb=None, quelle: str = "") -> Workbook:
    """Export aus der gerechneten Mappe wb (Werte); formeln_wb (dieselbe Mappe mit Formeln)
    liefert die Herkunft der Objektdaten, ohne sie bleiben die Zellen ohne Farbe."""
    ex = Workbook()
    spalten = _bwa_spalten(wb)
    blaetter = [UEBERBLICK, JAHRESRECHNUNG]
    _blatt_jahresrechnung(ex, wb)
    if _blatt_bruecke(ex, wb):
        blaetter.append(BRUECKE)
    blaetter.append(BWA_GESAMT)
    _blatt_bwa_gesamt(ex, wb, spalten)
    if _blatt_bwa_halten(ex, wb, spalten):
        blaetter.append(BWA_HALTEN)
    blaetter.append(BWA_OBJEKTE)
    _blatt_bwa_objekte(ex, wb, spalten)
    if _blatt_verkauf_kauf(ex, wb):
        blaetter.append(VERKAUF_KAUF)
    _blatt_annahmen(ex, wb, formeln_wb)
    blaetter.append(ANNAHMEN)
    _blatt_ueberblick(ex, wb, quelle, blaetter)
    ex.active = 0
    for ws in ex.worksheets:   # nur Werte: Texte wie „= Ergebnis“ sind keine Formeln
        for zeile in ws.iter_rows():
            for c in zeile:
                if c.data_type == "f":
                    c.data_type = "s"
    return ex


def _soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def lade_werte(pfad: Path, arbeitsordner: Path):
    """Mappe mit Werten laden; nie durchgerechnet: per LibreOffice rechnen, falls vorhanden."""
    wb = load_workbook(pfad, data_only=True)
    if "aus_VermoegenNetto" not in wb.defined_names:
        sys.exit(f"{pfad} ist keine Mappe des Prognosemodells.")
    if _wert(wb, "aus_VermoegenNetto") is not None:
        return wb
    soffice = _soffice()
    if not soffice:
        sys.exit(f"{pfad} enthält keine berechneten Werte. Bitte in Excel öffnen, speichern "
                 "und den Export erneut starten.")
    print("Die Mappe enthält keine berechneten Werte, LibreOffice rechnet sie durch.")
    kopie = arbeitsordner / f"mappe{pfad.suffix}"
    shutil.copyfile(pfad, kopie)
    aus = arbeitsordner / "gerechnet"
    subprocess.run(
        [soffice, f"-env:UserInstallation={(arbeitsordner / 'lo-profil').as_uri()}",
         "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(aus), str(kopie)],
        check=True, capture_output=True, timeout=600)
    return load_workbook(aus / "mappe.xlsx", data_only=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("mappe", help="gerechnete Mappe (.xlsx oder .xlsm)")
    ap.add_argument("--ausgabe", metavar="PFAD",
                    help="Exportdatei; Standard: <Mappe>_Mandant.xlsx daneben")
    args = ap.parse_args()
    quelle = Path(args.mappe)
    if not quelle.is_file():
        sys.exit(f"{quelle} nicht gefunden.")
    name = f"{quelle.stem}_Mandant.xlsx"
    ziel = Path(args.ausgabe) if args.ausgabe else quelle.with_name(name)
    if ziel.is_dir():
        ziel = ziel / name
    ziel = ziel.with_suffix(".xlsx")
    if ziel.resolve() == quelle.resolve():
        sys.exit("Die Mappe würde überschrieben: bitte mit --ausgabe einen anderen Namen "
                 "angeben.")
    with tempfile.TemporaryDirectory() as tmp:
        wb = lade_werte(quelle, Path(tmp))
        formeln_wb = load_workbook(quelle, read_only=True)
        try:
            ex = erstelle_export(wb, formeln_wb, quelle.name)
        finally:
            formeln_wb.close()
    ziel.parent.mkdir(parents=True, exist_ok=True)
    try:
        ex.save(ziel)
    except PermissionError:
        sys.exit(f"Kann {ziel} nicht schreiben. Ist die Datei noch in Excel geöffnet?")
    print(f"geschrieben: {ziel} ({', '.join(ex.sheetnames)})")


if __name__ == "__main__":
    main()
