"""Prüfskript: Mappe generieren, mit LibreOffice headless durchrechnen, gegen Sollwerte prüfen.

Aufruf: python -m pruefung.pruefen [-j PROZESSE] [Teil des Fallnamens]
Beendet sich mit Fehlercode 1, sobald ein Fall abweicht.

Die Fälle laufen parallel (Standard: ein Prozess je Kern) und rechnen mit einer kleinen
Mappe (TEST_KAPAZITAET), denn jede leere Zeile kostet LibreOffice Rechenzeit. Die Fälle in
VOLLE_GROESSE rechnen mit der Mappe in Originalgröße.

Fälle mit Makroaufrufen (Etappe 9) werden als .xlsm gebaut, in LibreOffice mit
Makros geöffnet, die Makros ausgeführt und das Ergebnis danach geprüft.
"""

import argparse
import concurrent.futures
import contextlib
import dataclasses
import io
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from prognosemodell.einlesen import (anlagen_zusammenfuehren, lese_inventar, lese_kostenstellen,
                                     ordne_anlagen_zu, zusammenfuehren)
from prognosemodell.makros import LibreOffice, speichere_mit_makros
from prognosemodell.mappe import erstelle_mappe
from prognosemodell.modelle import (FEHLER, HINWEIS, PRUEFUNGEN, Anlage, Kapazitaet,
                                    LIQ_SPALTEN, NEU_FELDER, NEU_SPALTEN, VERKAUF_FELDER,
                                    VERKAUF_SPALTEN,
                                    STATUS_ANNAHME_GELOESCHT, WARNUNG, Modell, Neuobjekt,
                                    Objekt, Verkauf, prognosejahre)
from openpyxl.utils import get_column_letter

from prognosemodell.bwa import ZEILEN
from prognosemodell.formeln import NK_ERSTE, nk_spalten, spalte
from prognosemodell.vorlagen import SPALTE_PLAN
from prognosemodell.testdaten import testobjekt
from prognosemodell.vorlagen import erstelle_inventar_vorlage, erstelle_vorlage

TOLERANZ = 0.01  # ein Cent
# Fälle vor der Alterslogik: Sollwerte ohne Alterung, Anlaufminderung und Großmaßnahmen.
# Gilt für jeden Fall, dessen Name nicht mit „Erhaltung“ beginnt.
OHNE_ALTERUNG = {"par_ErhAlterung": 0, "par_NeuErhAnlaufFaktor": 1, "par_SanQuote": 0}
# Zeilen je Eingabeblatt in den Prüfmappen: reicht für jeden Fall, rechnet viel schneller
TEST_KAPAZITAET = Kapazitaet(objekte=10, verkaeufe=10, neuobjekte=10, neukauf=5)
# Fälle, die zur Kontrolle mit der Mappe in Originalgröße rechnen (ohne prg_neu)
VOLLE_GROESSE = ("Etappe 7: Abnahme Summen über alle Objekte, Plan gleich Baseline",
                 "Etappe 9: Testobjekt ohne Befund")
KAPITALANLAGE = "Kapitalanlage (Nettoerlös − Reinvestition + Kredit − Steuer ca. − Ablösung)"
FEHLT = object()        # Makro ohne Rückgabewert: nur prüfen, dass es fehlerfrei läuft
AUSGEBLENDET = "Zeile ausgeblendet"  # statt Bereichsname: Zeilennummer im Blatt Prognose
EINGEKLAPPT = "Spalte eingeklappt:"  # + Blattname; statt Zeile: Spaltennummer, Soll (Ebene, aus)
ZEILE_EINGEKLAPPT = "Zeile eingeklappt:"  # + Blattname; statt Zeile: Zeilennummer, Soll (Ebene, aus)
KOPF = "Kopf:"  # + Blatt:Zeile; statt Zeile: Spaltennummer, Soll Text, " [Tooltip]" mit Kommentar
AUSWAHL = "Auswahl:"  # + Blattname; statt Zeile: Zelle, Soll (Liste, Anzahl Überprüfungen)
UEBERLAPPT = "Datenüberprüfungen überlappen"  # Zeile 0, Soll: Liste der Fundstellen ([])


def nr(spalten, key: str, versatz: int = 0) -> int:
    """Spaltennummer des Schlüssels key in einer Spaltenliste."""
    return versatz + 1 + [s.key for s in spalten].index(key)


@dataclasses.dataclass(frozen=True)
class Wie:
    """Sollwert = Wert eines anderen Bereichs derselben Mappe."""
    name: str
    zeile: int = 0


@dataclasses.dataclass
class Getippt:
    """Modell, dazu nach dem Generieren eingetippte Zellen {(Blatt, Zelle): Wert}."""
    modell: Modell
    zellen: dict


def mappe(modell: Modell, zellen: dict = None):
    wb = erstelle_mappe(modell)
    for (blatt, zelle), wert in (zellen or {}).items():
        wb[blatt][zelle] = wert
    return wb


def durchrechnen(modell: Modell, arbeitsordner: Path, zellen: dict = None, profil: Path = None):
    """Mappe schreiben, per LibreOffice neu berechnen lassen, Werte zurückgeben.

    Ein gemeinsames Profil für mehrere Aufrufe spart den Aufbau bei jedem Start."""
    roh = arbeitsordner / "roh" / "mappe.xlsx"
    roh.parent.mkdir(parents=True, exist_ok=True)
    erzeugt = mappe(modell, zellen)
    erzeugt.save(roh)
    aus = arbeitsordner / "gerechnet"
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("LibreOffice (soffice) nicht gefunden.")
    subprocess.run(
        [soffice, f"-env:UserInstallation={(profil or arbeitsordner / 'lo-profil').as_uri()}", "--headless",
         "--calc", "--convert-to", "xlsx", "--outdir", str(aus), str(roh)],
        check=True, capture_output=True, timeout=120,
    )
    wb = load_workbook(aus / "mappe.xlsx", data_only=True)
    wb.erzeugt = erzeugt  # Datenüberprüfungen so, wie Excel sie öffnet (LibreOffice glättet sie)
    return wb


def wert(wb, name: str, zeile_im_bereich: int = 0):
    """Wert eines benannten Bereichs (Spaltenbereich: n-te Zeile, Zeilenbereich: n-te Spalte).

    Statt eines Namens gehen auch "BWA:<Blatt>:<BWA-Nr.>" (n = Jahr),
    "SB<Vorgang>:<Beschriftung>" und "KAUF:<Beschriftung>" (n = Spalte, 1 = B).
    """
    if name.startswith("BWA:"):
        _, blatt, nr = name.split(":")
        return bwa_wert(wb[blatt], int(nr), zeile_im_bereich)
    if name.startswith(("SB", "KAUF:")):
        kennung, beschriftung = name.split(":", 1)
        start = "Detailsicht Kauf" if kennung == "KAUF" else f"Vorgang {kennung[2:]}:"
        return sonderbereich(wb, start, beschriftung, zeile_im_bereich)
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    teile = ref.replace("$", "").split(":")
    zelle = wb[blatt][teile[0]]
    if len(teile) == 2 and wb[blatt][teile[1]].row == zelle.row:
        return wb[blatt].cell(row=zelle.row, column=zelle.column + zeile_im_bereich).value
    return wb[blatt].cell(row=zelle.row + zeile_im_bereich, column=zelle.column).value


def bwa_wert(ws, nr: int, jahr: int):
    """Wert einer BWA-Zeile im Jahr: 2024/2025 Ist-Vorjahre, 2026 Ist, ab 2027 Plan."""
    from prognosemodell.vorlagen import SPALTE_JAHR, SPALTE_PLAN, SPALTE_VORJAHRE
    spalte = (SPALTE_JAHR if jahr == 2026 else SPALTE_PLAN + jahr - 2027 if jahr > 2026
              else SPALTE_VORJAHRE + jahr - 2024)
    for zeile in range(6, ws.max_row + 1):
        if ws.cell(row=zeile, column=2).value == nr:
            return ws.cell(row=zeile, column=spalte).value
    raise KeyError(f"BWA {nr} fehlt in Blatt {ws.title}")


def sonderbereich(wb, start: str, beschriftung: str, spalte: int):
    """Wert im Blatt Verkauf und Kauf: erste Zeile mit der Beschriftung nach dem Blocktitel."""
    ws = wb["Verkauf und Kauf"]
    im_block = False
    for zeile in range(1, ws.max_row + 1):
        text = ws.cell(row=zeile, column=1).value
        if isinstance(text, str) and text.startswith(start):
            im_block = True
            if text.startswith(beschriftung):
                return text
        elif im_block and text == beschriftung:
            return ws.cell(row=zeile, column=1 + spalte).value
    raise KeyError(f"{start} / {beschriftung} nicht im Blatt Verkauf und Kauf")


def gleich(ist, soll) -> bool:
    if isinstance(soll, (int, float)) and isinstance(ist, (int, float)):
        return abs(ist - soll) <= TOLERANZ
    return ist == soll


def prg(objekt_nr: int, jahr: int) -> int:
    """Zeile im Prognosebereich für Objekt objekt_nr (1 = erstes) und Jahr."""
    return (objekt_nr - 1) * prognosejahre() + jahr - 2027


def prg_neu(neu_nr: int, jahr: int) -> int:
    """Zeile im Prognosebereich für Neuobjekt neu_nr; die Neuobjekte folgen auf alle Objektblöcke."""
    return prg(TEST_KAPAZITAET.objekte + neu_nr, jahr)


def rls(jahr: int) -> int:
    """Zeile im Rücklagenspiegel je Jahr; erste Zeile ist das erste Prognosejahr."""
    return jahr - 2027


def lj(jahr: int) -> int:
    """Zeile in den Jahrestabellen von Liquidität und Auswertung; erste Zeile ist 2027."""
    return jahr - 2027


def pr(key: str) -> int:
    """Zeile einer Prüfung im Blatt Prüfung."""
    return [p.key for p in PRUEFUNGEN].index(key)


def befund(**anzahl) -> list:
    """Sollwerte aller Prüfzeilen: genannte Anzahl, sonst 0; dazu das Gesamtergebnis."""
    soll = []
    fehler = warnungen = 0
    for p in PRUEFUNGEN:
        n = anzahl.pop(p.key, 0)
        soll += [("pr_Anzahl", pr(p.key), n), ("pr_Ergebnis", pr(p.key), p.art if n else "OK")]
        fehler += bool(n) and p.art == FEHLER
        warnungen += bool(n) and p.art == WARNUNG
    assert not anzahl, f"unbekannte Prüfungen: {anzahl}"
    gesamt = (f"{fehler} {FEHLER}" if fehler else
              f"{warnungen} Warnung(en)" if warnungen else "OK")
    return soll + [("pr_Fehler", 0, fehler), ("pr_Warnungen", 0, warnungen),
                   ("pr_Gesamt", 0, gesamt)]


def ueb(jahr: int) -> int:
    """Zeile im Übersichtsbereich; erste Zeile ist das Basisjahr."""
    return jahr - 2026


# Je Fall: Name, Objekte (oder ein Modell mit Verkäufen), Liste von (benannter Bereich, Zeile im Bereich, Sollwert)
# Sollwerte der Etappe 2 von Hand: AfA 800.000 × 2 % = 16.000, Miete 60.000 × 1,02^n,
# Erhaltung 8.000 × 1,025^n, weitere Ausgaben × 1,02^n, n = Jahr − 2026.
# Übersicht: Verkehrswert 1.400.000 × 1,02^n je Objekt.
def vorlage_eingelesen(plan: dict = None, basis: dict = None, neukauf=(),
                       neuobjekte=()) -> Modell:
    """Modell aus der BWA-Vorlage: KSt 1 mit den Stammdaten des Testobjekts, KSt 2 ohne.

    plan: schon gefüllte Planspalten der Vorlage (erstelle_vorlage); basis: {BWA-Nr.: Wert}
    im Kostenstellenblatt KSt 1, Spalte Basisjahr, nach dem Einlesen geändert (wie im
    Blatt KSt 1 der Mappe getippt); neukauf: Neukauf-Kostenstellen hinter „KSt 9999“."""
    from prognosemodell.vorlagen import SPALTE_JAHR
    with tempfile.TemporaryDirectory() as tmp:
        pfad = Path(tmp) / "vorlage.xlsx"
        erstelle_vorlage(plan=plan, neukauf=neukauf).save(pfad)
        laufende, _ = lese_kostenstellen(pfad, 2026)
    nk = {lw.objekt_id: lw for lw in laufende if lw.neukauf}
    laufende = [lw for lw in laufende if not lw.neukauf]
    for nr, wert_ in (basis or {}).items():
        laufende[0].ist.setdefault(nr, {})[SPALTE_JAHR] = wert_
    stamm = [dataclasses.replace(testobjekt(), objekt_id="KSt 1", name=None)]
    return Modell(objekte=zusammenfuehren(stamm, laufende), neukauf=nk,
                  kostenstellen={lw.objekt_id: lw for lw in laufende},
                  neuobjekte=list(neuobjekte))


def anlagen_eingelesen(inventar: bool = True, ohne_kost1=(), **parameter) -> Modell:
    """Modell nur aus den Vorlagen: BWA je Kostenstelle, Anlagenverzeichnis (KSt 1 = Daten des
    Testobjekts, KSt 2 mit auslaufender Außenanlage und Anlage im Bau). Ohne Inventar zählt
    die Aufschlüsselung der Abschreibungen aus den Kostenstellenblättern. ohne_kost1: weitere
    Anlagen ohne KOST1, zugeordnet über die Bezeichnung."""
    with tempfile.TemporaryDirectory() as tmp:
        bwa, inv = Path(tmp) / "bwa.xlsx", Path(tmp) / "inventar.xlsx"
        erstelle_vorlage().save(bwa)
        erstelle_inventar_vorlage().save(inv)
        laufende, _ = lese_kostenstellen(bwa, 2026)
        anlagen = lese_inventar(inv)[0] if inventar else []
    objekte = zusammenfuehren([], laufende)
    anlagen = ordne_anlagen_zu(anlagen + list(ohne_kost1), objekte)
    return Modell(objekte=objekte, kostenstellen={lw.objekt_id: lw for lw in laufende},
                  anlagen=anlagen_zusammenfuehren(anlagen, laufende)[0], parameter=parameter)


def baujahr_bestaetigt() -> Modell:
    """Wie anlagen_eingelesen, aber für KSt 1 ein eingetipptes Baujahr (Bestandsgebäude)."""
    modell = anlagen_eingelesen()
    modell.objekte[0] = dataclasses.replace(modell.objekte[0], baujahr=1990)
    return modell


def faelle():
    obj = testobjekt()
    ohne_miete = dataclasses.replace(obj, miete=None)
    zu_hoch = dataclasses.replace(obj, restbuchwert=900_000)
    spaet = dataclasses.replace(obj, kaufjahr=2030)
    zweites = dataclasses.replace(obj, name="Duplikat")
    laeuft_aus = dataclasses.replace(obj, restbuchwert=50_000)
    satz_25 = dataclasses.replace(obj, afa_satz=0.025, restbuchwert=400_000)
    obj2 = dataclasses.replace(obj, objekt_id="OBJ-002", weitere_einnahmen=1_000,
                               weitere_ausgaben=2_000)
    obj3 = dataclasses.replace(obj, objekt_id="OBJ-003", verkehrswert=None)
    # Etappe 7 rechnet ohne Zins auf die Liquidität
    OHNE_ZINS = {"par_Alternativrendite": 0}
    abnahme4 = dataclasses.replace(obj, restbuchwert=496_000)  # Buchwert Ende 2027: 480.000
    # "" = Annahme in der Zelle gelöscht, None = Annahme greift
    ohne_quote = dataclasses.replace(obj, objekt_id="OBJ-005", vk_quote_gebaeude="")
    jung = dataclasses.replace(obj, objekt_id="OBJ-004", kaufjahr=2024)
    return [
        ("Etappe 1: Stammdaten vollständig", [obj], [
            ("par_Startjahr", 0, 2027),
            ("par_Endjahr", 0, 2046),
            ("obj_ID", 0, "OBJ-001"),
            ("obj_Status", 0, "OK"),
            ("obj_Status", 1, None),  # leere Zeile bleibt leer
        ]),
        ("Etappe 1: Pflichtfeld fehlt", [ohne_miete], [("obj_Status", 0, "Pflichtfeld fehlt")]),
        ("Etappe 1: Restbuchwert zu hoch", [zu_hoch],
         [("obj_Status", 0, "Restbuchwert > AK Gebäude")]),
        ("Etappe 1: Kaufjahr nach Basisjahr", [spaet],
         [("obj_Status", 0, "Kaufjahr nach Basisjahr")]),
        ("Etappe 1: ObjektID doppelt", [obj, zweites],
         [("obj_Status", 0, "ObjektID doppelt"), ("obj_Status", 1, "ObjektID doppelt")]),
        ("Etappe 2: AfA 2 %, Kauf 2007, Restbuchwert 480.000", [obj], [
            ("prg_ID", prg(1, 2027), "OBJ-001"),
            ("prg_Jahr", prg(1, 2027), 2027),
            ("prg_Jahr", prg(1, 2046), 2046),
            ("prg_AfA", prg(1, 2027), 16_000),
            ("prg_Buchwert", prg(1, 2027), 464_000),
            ("prg_Miete", prg(1, 2027), 61_200),
            ("prg_Erhaltung", prg(1, 2027), 8_200),
            ("prg_Ergebnis", prg(1, 2027), 37_000),
            ("prg_AfA", prg(1, 2046), 16_000),
            ("prg_Buchwert", prg(1, 2046), 160_000),  # Nullpunkt erst Ende 2056
            ("prg_Miete", prg(1, 2046), 89_156.84),
            ("prg_Ergebnis", prg(1, 2046), 60_047.91),
            ("prg_ID", prg(2, 2027), None),  # leere Objektzeile, leerer Block
            ("prg_Ergebnis", prg(2, 2027), None),
        ]),
        ("Etappe 2: Buchwert läuft im Raster aus (Restbuchwert 50.000)", [laeuft_aus], [
            ("prg_AfA", prg(1, 2029), 16_000),
            ("prg_Buchwert", prg(1, 2029), 2_000),
            ("prg_AfA", prg(1, 2030), 2_000),  # nur noch der Rest
            ("prg_Buchwert", prg(1, 2030), 0),
            ("prg_AfA", prg(1, 2031), 0),
            ("prg_Buchwert", prg(1, 2031), 0),
            ("prg_AfA", prg(1, 2046), 0),
            ("prg_Miete", prg(1, 2031), 66_244.85),  # Miete läuft weiter
            ("prg_Ergebnis", prg(1, 2031), 57_193.58),
        ]),
        ("Etappe 2: AfA 2,5 %, Kauf 2007, Nullpunkt Ende 2046", [satz_25], [
            ("prg_AfA", prg(1, 2027), 20_000),
            ("prg_Buchwert", prg(1, 2045), 20_000),
            ("prg_AfA", prg(1, 2046), 20_000),
            ("prg_Buchwert", prg(1, 2046), 0),
        ]),
        ("Etappe 2: zweites Objekt mit weiteren Einnahmen und Ausgaben", [obj, obj2], [
            ("prg_Einnahmen", prg(1, 2027), 0),
            ("prg_ID", prg(2, 2027), "OBJ-002"),
            ("prg_Jahr", prg(2, 2027), 2027),
            ("prg_Einnahmen", prg(2, 2027), 1_020),
            ("prg_Ausgaben", prg(2, 2027), 2_040),
            ("prg_Ergebnis", prg(2, 2027), 35_980),
        ]),
        ("Übersicht: ohne Verkauf sind Plan und Baseline gleich", [obj], [
            ("prg_Aktiv", prg(1, 2046), 1),
            ("prg_Bestand", prg(1, 2046), 1),
            ("prg_Verkehrswert", prg(1, 2027), 1_428_000),
            ("prg_Verkehrswert", prg(1, 2046), 2_080_326.35),
            ("ueb_Jahr", ueb(2026), 2026),
            ("ueb_Jahr", ueb(2046), 2046),
            ("ueb_Baseline", ueb(2026), 1_400_000),
            ("ueb_Plan", ueb(2026), 1_400_000),
            ("ueb_Baseline", ueb(2046), 2_080_326.35),
            ("ueb_Plan", ueb(2046), 2_080_326.35),
            ("ueb_Differenz", ueb(2046), 0),
            ("ueb_Objekte", 0, 1),
            ("ueb_MitAnnahmen", 0, 0),
            ("ueb_Verkaeufe", 0, 0),
        ]),
        ("Verkauf OBJ-001 Ende 2030: Plan verliert den Wert, Baseline hält",
         Modell(objekte=[obj, obj2], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_500_000)]), [
            ("vk_Status", 0, "OK"),
            ("prg_Aktiv", prg(1, 2030), 1),  # Miete und AfA laufen im Verkaufsjahr noch
            ("prg_Miete", prg(1, 2030), 64_945.93),
            ("prg_Aktiv", prg(1, 2031), 0),
            ("prg_Miete", prg(1, 2031), 0),
            ("prg_AfA", prg(1, 2031), 0),
            ("prg_Ergebnis", prg(1, 2031), 0),
            ("prg_Buchwert", prg(1, 2030), 416_000),
            ("prg_Bestand", prg(1, 2029), 1),
            ("prg_Bestand", prg(1, 2030), 0),  # Verkauf zum Jahresende
            ("prg_Aktiv", prg(2, 2046), 1),  # OBJ-002 bleibt
            ("prg_Bestand", prg(2, 2046), 1),
            ("ueb_Baseline", ueb(2026), 2_800_000),
            ("ueb_Plan", ueb(2026), 2_800_000),
            ("ueb_Plan", ueb(2029), 2_971_382.40),
            ("ueb_Baseline", ueb(2030), 3_030_810.04),
            ("ueb_Plan", ueb(2030), 1_515_405.02),
            ("ueb_Differenz", ueb(2030), -1_515_405.02),
            ("ueb_Baseline", ueb(2046), 4_160_652.70),
            ("ueb_Plan", ueb(2046), 2_080_326.35),
            ("ueb_Verkaeufe", 0, 1),
        ]),
        # OBJ-003 ohne Verkehrswert: Annahme 60.000 × 20 = 1,2 Mio
        ("Verkäufe: Statusprüfung, fehlender Verkehrswert angenommen",
         Modell(objekte=[obj, obj2, obj3],
                verkaeufe=[Verkauf("OBJ-001", 2030), Verkauf("OBJ-001", 2031),
                           Verkauf("XYZ", 2030), Verkauf("OBJ-002", 2050),
                           Verkauf("OBJ-003", None)]), [
            ("vk_Status", 0, "Objekt mehrfach verkauft"),
            ("vk_Status", 1, "Objekt mehrfach verkauft"),
            ("vk_Status", 2, "ObjektID unbekannt"),
            ("vk_Status", 3, "Verkaufsjahr außerhalb Raster"),
            ("vk_Status", 4, "Verkaufsjahr fehlt"),
            ("vk_Status", 5, None),
            ("prg_Aktiv", prg(3, 2046), 1),  # ohne Verkaufsjahr kein Verkauf
            ("ueb_MitAnnahmen", 0, 1),
            ("obj_Verkehrswert", 2, 1_200_000),
            ("obj_Annahmen", 2, 1),
            ("ueb_Baseline", ueb(2026), 4_000_000),
        ]),
        ("Etappe 4: Abnahme Preis 1,4 Mio, Buchwert gesamt 680.000, hälftig",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")]), [
            ("vk_Status", 0, "OK"),
            ("vk_Vorbesitz", 0, 20),
            ("vk_BuchwertGeb", 0, 480_000),
            ("vk_AKGuB", 0, 200_000),
            ("vk_Nettoerloes", 0, 1_400_000),
            ("vk_AnteilGuB", 0, 0.5),  # aus Verkehrswertanteil Gebäude 50 %
            ("vk_ErloesGeb", 0, 700_000),
            ("vk_ErloesGuB", 0, 700_000),
            ("vk_GewinnGeb", 0, 220_000),
            ("vk_GewinnGuB", 0, 500_000),
            ("vk_Gewinn", 0, 720_000),
            ("vk_Gewinn", 1, None),  # leere Zeile bleibt leer
        ]),
        ("Etappe 4: Kaufvertrag 30 % G+B, Kosten 40.000, Verkauf Ende 2030",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000,
                                                  kosten=40_000, anteil_gub=0.3)]), [
            ("vk_Status", 0, "OK"),
            ("vk_BuchwertGeb", 0, 416_000),  # 480.000 − 4 × 16.000, AfA 2030 noch enthalten
            ("vk_Nettoerloes", 0, 1_360_000),
            ("vk_AnteilGuB", 0, 0.3),  # Kaufvertrag vor Verkehrswert
            ("vk_ErloesGuB", 0, 408_000),
            ("vk_ErloesGeb", 0, 952_000),
            ("vk_GewinnGeb", 0, 536_000),
            ("vk_GewinnGuB", 0, 208_000),
            ("vk_Gewinn", 0, 744_000),
        ]),
        ("Etappe 4: Statusprüfung Preis, Aufteilung, § 6b-Vorbesitzzeit",
         Modell(objekte=[obj, jung, ohne_quote],
                verkaeufe=[Verkauf("OBJ-001", 2028, preis=""),
                           Verkauf("OBJ-005", 2028, preis=1_000_000),
                           Verkauf("OBJ-004", 2029, preis=1_000_000, nutzung_6b="ja")]), [
            ("vk_Status", 0, "Verkaufspreis fehlt"),
            ("vk_Status", 1, "Aufteilung fehlt: Anteil G+B oder Verkehrswertanteil"),
            ("obj_Status", 2, STATUS_ANNAHME_GELOESCHT),
            ("vk_Gewinn", 1, None),
            ("vk_Vorbesitz", 2, 5),
            ("vk_Status", 2, "§ 6b unzulässig: Vorbesitzzeit zu kurz"),
            ("rl_ID", 2, None),  # keine Rücklage, Gewinn wird sofort steuerwirksam
            # Buchwert Ende 2029 432.000, Erlös je 500.000: 68.000 + 300.000
            ("rls_Gewinne", rls(2029), 368_000),
            ("rls_Bildung", rls(2029), 0),
            ("rls_Steuerwirksam", rls(2029), 368_000),
            ("rls_Gewinne", rls(2028), 0),  # Zeilen ohne Status OK zählen nicht
        ]),
        # Etappe 5: Gewinn 720.000 aus dem Abnahmefall der Etappe 4, Fristjahr 2027 + 4
        ("Etappe 5: Abnahme Rücklage voll, Auflösung mit Zuschlag im Fristjahr",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")]), [
            ("par_6bFristNeubau", 0, 6),
            ("rl_ID", 0, "RL-OBJ-001"),
            ("rl_ObjektID", 0, "OBJ-001"),
            ("rl_Jahr", 0, 2027),
            ("rl_Geb", 0, 220_000),
            ("rl_GuB", 0, 500_000),
            ("rl_Betrag", 0, 720_000),
            ("rl_Fristjahr", 0, 2031),
            ("rl_Aufloesung", 0, 720_000),
            ("rl_Zuschlag", 0, 172_800),  # 720.000 × 6 % × 4
            ("rl_Hinweis", 0, None),
            ("rl_ID", 1, None),
            ("rls_Jahr", rls(2027), 2027),
            ("rls_Jahr", rls(2046), 2046),
            ("rls_Gewinne", rls(2027), 720_000),
            ("rls_Bildung", rls(2027), 720_000),
            ("rls_Steuerwirksam", rls(2027), 0),  # Steuer im Verkaufsjahr null
            ("rls_Bestand", rls(2027), 720_000),
            ("rls_Bestand", rls(2030), 720_000),
            ("rls_Steuerwirksam", rls(2030), 0),
            ("rls_Aufloesung", rls(2031), 720_000),
            ("rls_Zuschlag", rls(2031), 172_800),
            ("rls_Steuerwirksam", rls(2031), 892_800),
            ("rls_Bestand", rls(2031), 0),
            ("rls_Bestand", rls(2046), 0),
        ]),
        ("Etappe 5: Neubau begonnen, Frist sechs Jahre",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja",
                                   neubau_6b="ja")]), [
            ("vk_Status", 0, "OK"),
            ("rl_Fristjahr", 0, 2033),
            ("rl_Zuschlag", 0, 259_200),  # 720.000 × 6 % × 6
            ("rls_Aufloesung", rls(2031), 0),
            ("rls_Bestand", rls(2032), 720_000),
            ("rls_Steuerwirksam", rls(2033), 979_200),
            ("rls_Bestand", rls(2033), 0),
        ]),
        ("Etappe 5: § 6b nein, Gewinn sofort steuerwirksam",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000,
                                                  kosten=40_000, anteil_gub=0.3,
                                                  nutzung_6b="nein")]), [
            ("rl_ID", 0, None),
            ("rl_Betrag", 0, None),
            ("rls_Gewinne", rls(2030), 744_000),
            ("rls_Bildung", rls(2030), 0),
            ("rls_Steuerwirksam", rls(2030), 744_000),
            ("rls_Bestand", rls(2030), 0),
        ]),
        # Kaufvertrag 80 % G+B: Gebäude 280.000 − 416.000, G+B 1.120.000 − 200.000
        ("Etappe 5: Verlust Gebäude mindert die Rücklage G+B nicht",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000,
                                                  anteil_gub=0.8, nutzung_6b="ja")]), [
            ("vk_GewinnGeb", 0, -136_000),
            ("vk_GewinnGuB", 0, 920_000),
            ("rl_Geb", 0, 0),
            ("rl_GuB", 0, 920_000),
            ("rl_Betrag", 0, 920_000),
            ("rls_Gewinne", rls(2030), 784_000),
            ("rls_Bildung", rls(2030), 920_000),
            ("rls_Steuerwirksam", rls(2030), -136_000),  # Verlust wirkt sofort
            ("rls_Bestand", rls(2030), 920_000),
        ]),
        # Buchwert Ende 2044: 480.000 − 18 × 16.000 = 192.000; Gewinn 508.000 + 500.000
        ("Etappe 5: Frist endet nach Prognoseende",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2044, preis=1_400_000,
                                                  nutzung_6b="ja")]), [
            ("rl_Betrag", 0, 1_008_000),
            ("rl_Fristjahr", 0, 2048),
            ("rl_Hinweis", 0, "Frist endet nach Prognoseende"),
            ("rls_Bestand", rls(2046), 1_008_000),
            ("rls_Aufloesung", rls(2046), 0),
        ]),
        # Etappe 6: Rücklage aus dem Abnahmefall (Gebäude 220.000, G+B 500.000), Kauf Ende 2028
        ("Etappe 6: Abnahme Übertragung, AfA-Basis 480.000, AK G+B 0",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")]), [
            ("ne_Status", 0, "OK"),
            ("ne_Gueltig", 0, 1),
            ("ne_AKGuBNeu", 0, 360_000),
            ("ne_AKGebNeu", 0, 840_000),
            ("ne_RLGeb", 0, 220_000),
            ("ne_RLGuB", 0, 500_000),
            ("ne_Ue1", 0, 220_000),
            ("ne_Ue2", 0, 360_000),
            ("ne_Ue3", 0, 140_000),
            ("ne_UeGesamt", 0, 720_000),
            ("ne_AfABasis", 0, 480_000),
            ("ne_AKGuB", 0, 0),
            ("ne_Status", 1, None),
            ("rl_UebGeb", 0, 220_000),
            ("rl_UebGuB", 0, 500_000),
            ("rl_Aufloesung", 0, 0),
            ("rl_Zuschlag", 0, 0),
            ("rls_Bestand", rls(2027), 720_000),
            ("rls_Uebertragung", rls(2028), 720_000),
            ("rls_Bestand", rls(2028), 0),
            ("rls_Steuerwirksam", rls(2028), 0),
            ("rls_Steuerwirksam", rls(2031), 0),  # nichts mehr aufzulösen
            # Prognose: Bestand ab Ende 2028, Miete und AfA ab 2029, AfA 3 % von 480.000
            ("prg_ID", prg_neu(1, 2027), "NEU-001"),
            ("prg_Neu", prg_neu(1, 2027), 1),
            ("prg_Neu", prg(1, 2027), 0),
            ("prg_Bestand", prg_neu(1, 2027), 0),
            ("prg_Bestand", prg_neu(1, 2028), 1),
            ("prg_Aktiv", prg_neu(1, 2028), 0),
            ("prg_Miete", prg_neu(1, 2028), 0),
            ("prg_Buchwert", prg_neu(1, 2027), 0),
            ("prg_Buchwert", prg_neu(1, 2028), 480_000),
            ("prg_AfA", prg_neu(1, 2028), 0),
            ("prg_Aktiv", prg_neu(1, 2029), 1),
            ("prg_AfA", prg_neu(1, 2029), 14_400),
            ("prg_Buchwert", prg_neu(1, 2029), 465_600),
            ("prg_Miete", prg_neu(1, 2029), 61_200),  # 1,2 Mio × 5 % × 1,02
            ("prg_Erhaltung", prg_neu(1, 2029), 12_300),  # 1,2 Mio × 1 % × 1,025
            ("prg_Ergebnis", prg_neu(1, 2029), 34_500),
            ("prg_Verkehrswert", prg_neu(1, 2028), 1_200_000),
            ("prg_Verkehrswert", prg_neu(1, 2029), 1_224_000),
            ("prg_Buchwert", prg_neu(1, 2046), 480_000 - 18 * 14_400),  # AfA 2029 bis 2046
            # Übersicht: Neuobjekt nur im Plan
            ("ueb_Neuobjekte", 0, 1),
            ("ueb_Plan", ueb(2027), 0),
            ("ueb_Baseline", ueb(2028), 1_456_560),
            ("ueb_Plan", ueb(2028), 1_200_000),
            ("ueb_Plan", ueb(2029), 1_224_000),
            ("prg_ID", prg_neu(2, 2027), None),
        ]),
        # NEU-001: G+B 100.000, Gebäude 400.000, nimmt 220.000 + 100.000 + 180.000.
        # NEU-002 mit Nebenkosten: AK 1.040.000, G+B 260.000, Gebäude 780.000,
        # bekommt den Rest G+B-Gewinn 500.000 − 100.000 − 180.000 = 220.000
        ("Etappe 6: zwei Neuobjekte teilen sich eine Rücklage",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[
                    Neuobjekt("NEU-001", 2028, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02,
                              quelle="RL-OBJ-001"),
                    Neuobjekt("NEU-002", 2030, kaufpreis=1_000_000, anteil_gub=0.25,
                              nebenkosten=40_000, afa_satz=0.02, quelle="RL-OBJ-001")]), [
            ("ne_Ue1", 0, 220_000),
            ("ne_Ue2", 0, 100_000),
            ("ne_Ue3", 0, 180_000),
            ("ne_AfABasis", 0, 0),
            ("ne_AKGuB", 0, 0),
            ("ne_AKGuBNeu", 1, 260_000),
            ("ne_AKGebNeu", 1, 780_000),
            ("ne_RLGeb", 1, 0),
            ("ne_RLGuB", 1, 220_000),
            ("ne_Ue1", 1, 0),
            ("ne_Ue2", 1, 220_000),
            ("ne_Ue3", 1, 0),
            ("ne_AfABasis", 1, 780_000),
            ("ne_AKGuB", 1, 40_000),
            ("rl_UebGeb", 0, 220_000),
            ("rl_UebGuB", 0, 500_000),
            ("rl_Aufloesung", 0, 0),
            ("rls_Uebertragung", rls(2028), 500_000),
            ("rls_Bestand", rls(2028), 220_000),
            ("rls_Uebertragung", rls(2030), 220_000),
            ("rls_Bestand", rls(2030), 0),
            ("prg_AfA", prg_neu(1, 2029), 0),  # AfA-Basis 0
            ("prg_AfA", prg_neu(2, 2031), 15_600),
        ]),
        # Neuobjekt G+B 150.000, Gebäude 150.000: Rest Gebäude 70.000 + G+B 350.000
        ("Etappe 6: Teilübertragung, Rest wird im Fristjahr aufgelöst",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2029, kaufpreis=300_000, anteil_gub=0.5,
                                      afa_satz=0.02, quelle="RL-OBJ-001")]), [
            ("ne_Ue1", 0, 150_000),
            ("ne_Ue2", 0, 150_000),
            ("ne_Ue3", 0, 0),
            ("ne_AfABasis", 0, 0),
            ("rl_UebGeb", 0, 150_000),
            ("rl_UebGuB", 0, 150_000),
            ("rl_Aufloesung", 0, 420_000),
            ("rl_Zuschlag", 0, 100_800),  # 420.000 × 6 % × 4
            ("rls_Bestand", rls(2029), 420_000),
            ("rls_Steuerwirksam", rls(2031), 520_800),
            ("rls_Bestand", rls(2031), 0),
        ]),
        # AK G+B von Hand fest 300.000, Gebäude = Rest 1,2 Mio − 300.000; NEU-002 mit
        # beiden Beträgen (Summe 50.000 über dem Kaufpreis): Prüfung warnt
        ("Etappe 6: AK G+B und AK Gebäude von Hand",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[
                    Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                              afa_satz=0.03, ak_gub_eingabe=300_000, quelle="RL-OBJ-001"),
                    Neuobjekt("NEU-002", 2028, kaufpreis=500_000, anteil_gub=0.2,
                              afa_satz=0.02, ak_gub_eingabe=150_000, ak_geb_eingabe=400_000),
                    Neuobjekt("NEU-003", 2028, kaufpreis=500_000, anteil_gub=0.2,
                              nebenkosten=50_000, afa_satz=0.02)]), [
            ("ne_AKGuBEingabe", 0, 300_000),
            ("ne_AKGebEingabe", 0, 900_000),   # vorbelegt: Rest nach G+B
            ("ne_AKGuBNeu", 0, 300_000),
            ("ne_AKGebNeu", 0, 900_000),
            ("ne_AKAbweichung", 0, 0),
            ("ne_Ue1", 0, 220_000),
            ("ne_Ue2", 0, 300_000),
            ("ne_Ue3", 0, 200_000),
            ("ne_AfABasis", 0, 480_000),
            ("ne_AKGuB", 0, 0),
            ("ne_AKGuBNeu", 1, 150_000),
            ("ne_AKGebNeu", 1, 400_000),
            ("ne_AKAbweichung", 1, 50_000),
            ("ne_Status", 1, "OK"),
            ("ne_AKGuBEingabe", 2, 110_000),   # (500.000 + 50.000) × 20 %
            ("ne_AKGebEingabe", 2, 440_000),
            ("ne_AKAbweichung", 2, 0),
            ("pr_Anzahl", pr("ak_aufteilung"), 1),
            ("pr_Ergebnis", pr("ak_aufteilung"), WARNUNG),
            ("ne_AKGuBEingabe", 3, None),
        ]),
        ("Etappe 6: Statusprüfung Neuobjekte",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[
                    Neuobjekt("NEU-001", 2032, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02,
                              quelle="RL-OBJ-001"),
                    Neuobjekt("NEU-002", 2028, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02,
                              quelle="RL-XYZ"),
                    Neuobjekt("OBJ-001", 2028, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02),
                    Neuobjekt("NEU-004", 2028, kaufpreis=500_000, anteil_gub=0.2),
                    Neuobjekt("NEU-005", 2050, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02),
                    Neuobjekt("NEU-006", 2028, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02),
                    Neuobjekt("NEU-006", 2029, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02),
                ]), [
            ("ne_Status", 0, "Kauf nach Fristjahr, keine Übertragung"),
            ("ne_Gueltig", 0, 1),  # bleibt im Modell, nur ohne Übertragung
            ("ne_Ue1", 0, 0),
            ("ne_AfABasis", 0, 400_000),
            ("prg_Bestand", prg_neu(1, 2032), 1),
            ("ne_Status", 1, "Rücklage unbekannt, keine Übertragung"),
            ("ne_Ue2", 1, 0),
            ("ne_Status", 2, "NeuID wie Bestandsobjekt"),
            ("ne_Gueltig", 2, 0),
            ("prg_Bestand", prg_neu(3, 2046), 0),
            ("vk_BuchwertGeb", 0, 480_000),  # Bestandsobjekt bleibt unverfälscht
            ("ne_Status", 3, "Pflichtfeld fehlt"),
            ("ne_Gueltig", 3, 0),
            ("ne_Status", 4, "Kaufjahr außerhalb Raster"),
            ("ne_Status", 5, "NeuID doppelt"),
            ("ne_Status", 6, "NeuID doppelt"),
            ("rl_Aufloesung", 0, 720_000),
            ("ueb_Neuobjekte", 0, 2),
        ]),
        # Verkauf Ende 2030, Neuobjekt Ende 2028: Rücklage gab es da noch nicht
        ("Etappe 6: Kauf vor Bildung der Rücklage",
         Modell(objekte=[obj],
                verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=500_000, anteil_gub=0.2,
                                      afa_satz=0.02, quelle="RL-OBJ-001")]), [
            ("ne_Status", 0, "Kauf vor Bildung der Rücklage, keine Übertragung"),
            ("ne_UeGesamt", 0, 0),
        ]),
        # AfA-Basis je 800.000 (Kaufpreis 1 Mio, G+B 20 %), Kauf Ende 2027, AfA ab 2028.
        # NEU-001 degressiv 5 % bei Nutzungsdauer 33⅓ Jahre (3 %): 2028 40.000, dann × 0,95;
        # 2042 Restnutzungsdauer 19⅓ < 20, Wechsel zu linear: 800.000 × 0,95^14 / 19⅓ = 20.179,65
        # und gleich bleibend. NEU-002 linear 3 % = 24.000. NEU-003 degressiv bei
        # Nutzungsdauer 20 Jahre (5 %): ab dem 2. Jahr ist linear höher, also stets 40.000.
        ("Etappe 6: degressive AfA mit Wechsel zur linearen AfA",
         Modell(objekte=[obj], neuobjekte=[
             Neuobjekt("NEU-001", 2027, kaufpreis=1_000_000, anteil_gub=0.2, afa_satz=0.03,
                       afa_methode="degressiv"),
             Neuobjekt("NEU-002", 2027, kaufpreis=1_000_000, anteil_gub=0.2, afa_satz=0.03),
             Neuobjekt("NEU-003", 2027, kaufpreis=1_000_000, anteil_gub=0.2, afa_satz=0.05,
                       afa_methode="degressiv")]), [
            ("prg_AfA", prg_neu(1, 2027), 0),
            ("prg_Buchwert", prg_neu(1, 2027), 800_000),
            ("prg_AfA", prg_neu(1, 2028), 40_000),
            ("prg_AfA", prg_neu(1, 2029), 38_000),
            ("prg_AfA", prg_neu(1, 2041), 20_533.68),
            ("prg_AfA", prg_neu(1, 2042), 20_179.65),
            ("prg_AfA", prg_neu(1, 2046), 20_179.65),
            ("prg_Buchwert", prg_neu(1, 2046), 289_241.71),
            ("prg_AfA", prg_neu(2, 2028), 24_000),
            ("prg_Buchwert", prg_neu(2, 2046), 344_000),
            ("prg_AfA", prg_neu(3, 2028), 40_000),
            ("prg_AfA", prg_neu(3, 2040), 40_000),
            ("prg_Buchwert", prg_neu(3, 2046), 40_000),
            ("ne_Status", 0, "OK"),
        ]),
        # Etappe 7: Steuersatz 30 %. Zwei Objekte 2027: Ergebnis 37.000 + 35.980,
        # Einnahmen 2 × 61.200 + 1.020, Ausgaben 2 × 8.200 + 2.040
        ("Etappe 7: Abnahme Summen über alle Objekte, Plan gleich Baseline",
         Modell(objekte=[obj, obj2], parameter=OHNE_ZINS), [
            ("liq_Jahr", lj(2027), 2027),
            ("liq_Jahr", lj(2046), 2046),
            ("liq_Ergebnis", lj(2027), 72_980),
            ("liq_Einnahmen", lj(2027), 123_420),
            ("liq_Ausgaben", lj(2027), 18_440),
            ("liq_Steuer", lj(2027), 21_894),
            ("liq_Zufluss", lj(2027), 83_086),
            ("liq_Kum", lj(2027), 83_086),
            ("liq_Ergebnis", lj(2046), 118_609.88),
            ("liq_Steuer", lj(2046), 35_582.96),
            ("liq_Kum", lj(2046), 1_963_197.65),
            ("lqb_Einnahmen", lj(2027), 123_420),
            ("lqb_AfA", lj(2027), 32_000),
            ("lqb_Ergebnis", lj(2027), 72_980),
            ("lqb_Ergebnis", lj(2046), 118_609.88),
            ("lqb_Kum", lj(2046), 1_963_197.65),
            ("aus_GuV", lj(2027), 72_980),
            ("aus_NachSteuer", lj(2027), 51_086),
            ("aus_SteuerKum", lj(2046), 567_084.71),
            # Verkehrswert 2 × 1.428.000, Buchwert 2 × (464.000 + 200.000)
            ("aus_Verkehrswert", lj(2027), 2_856_000),
            ("aus_Buchwert", lj(2027), 1_328_000),
            ("aus_StilleReserven", lj(2027), 1_528_000),
            ("aus_LatenteSteuer", lj(2027), 458_400),
            ("aus_Vermoegen", lj(2027), 2_939_086),
            ("aus_VermoegenNetto", lj(2027), 2_480_686),
            ("asb_VermoegenNetto", lj(2027), 2_480_686),
            ("aus_VermoegenNetto", lj(2046), 5_091_654.55),
            ("ueb_VermBaseline", ueb(2026), 2_368_000),  # 2,8 Mio − 1,44 Mio × 30 %
            ("ueb_VermPlan", ueb(2026), 2_368_000),
            ("ueb_VermPlan", ueb(2027), 2_480_686),
            ("ueb_VermBaseline", ueb(2046), 5_091_654.55),
            ("ueb_VermDifferenz", ueb(2046), 0),
        ]),
        # Gewinn 744.000 ohne § 6b; Ergebnis 2030 = 64.945,93 − 8.830,50 − 16.000
        ("Etappe 7: Abnahme Verkauf ohne § 6b, Steuer im Verkaufsjahr",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000,
                                                  kosten=40_000, anteil_gub=0.3,
                                                  nutzung_6b="nein")],
                parameter=OHNE_ZINS), [
            ("liq_Ergebnis", lj(2030), 40_115.43),
            ("liq_Verkauf", lj(2030), 744_000),
            ("liq_Steuer", lj(2030), 235_234.63),  # 12.034,63 + 744.000 × 30 %
            ("liq_Verkaufserloes", lj(2030), 1_360_000),
            ("liq_Rueckfluss", lj(2030), 616_000),  # Buchwert 416.000 + AK G+B 200.000
            ("liq_Kum", lj(2030), 1_308_734.25),
            ("liq_Ergebnis", lj(2031), 0),
            ("liq_Steuer", lj(2031), 0),
            ("liq_Zufluss", lj(2031), 0),
            ("liq_Kum", lj(2046), 1_308_734.25),
            ("lqb_Ergebnis", lj(2031), 41_193.58),  # Baseline hält das Objekt
            ("lqb_Kum", lj(2030), 171_934.25),
            ("aus_Verkehrswert", lj(2030), 0),
            ("aus_Buchwert", lj(2030), 0),
            ("aus_LatenteSteuer", lj(2030), 0),
            ("aus_VermoegenNetto", lj(2030), 1_308_734.25),
            ("asb_Verkehrswert", lj(2030), 1_515_405.02),
            ("asb_Buchwert", lj(2030), 616_000),
            ("asb_VermoegenNetto", lj(2030), 1_417_517.76),
            ("ueb_VermPlan", ueb(2030), 1_308_734.25),
            ("ueb_VermBaseline", ueb(2030), 1_417_517.76),
        ]),
        # Rücklage 720.000 ohne Reinvestition: Auflösung und Zuschlag 2031 werden versteuert
        ("Etappe 7: Rücklage ohne Reinvestition, Steuer im Fristjahr",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                parameter=OHNE_ZINS), [
            ("liq_Verkauf", lj(2027), 0),
            ("liq_Steuer", lj(2027), 11_100),  # nur laufendes Ergebnis 37.000
            ("liq_Verkaufserloes", lj(2027), 1_400_000),
            ("liq_Rueckfluss", lj(2027), 680_000),
            ("liq_Zufluss", lj(2027), 1_441_900),
            ("liq_Steuer", lj(2031), 267_840),  # 892.800 × 30 %
            ("liq_Kum", lj(2031), 1_441_900 - 267_840),
            ("aus_Ruecklage", lj(2027), 720_000),
            ("aus_LatenteSteuer", lj(2027), 216_000),  # Rücklage ist gestundete Steuer
            ("aus_VermoegenNetto", lj(2027), 1_225_900),
            ("aus_Ruecklage", lj(2031), 0),
            ("aus_LatenteSteuer", lj(2031), 0),
            ("aus_SteuerKum", lj(2031), 11_100 + 267_840),
        ]),
        # Kauf Ende 2028 für 1,2 Mio, AfA-Basis 480.000: Tausch Liquidität gegen Objekt
        ("Etappe 7: Reinvestition, Kauf aus der Liquidität",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")],
                parameter=OHNE_ZINS), [
            ("liq_Kauf", lj(2028), 1_200_000),
            ("liq_Zufluss", lj(2028), -1_200_000),
            ("liq_Kum", lj(2028), 241_900),
            ("liq_Ergebnis", lj(2029), 34_500),
            ("liq_Steuer", lj(2029), 10_350),
            ("liq_Zufluss", lj(2029), 38_550),  # 61.200 − 12.300 − 10.350
            ("liq_Steuer", lj(2031), 10_904.94),  # keine Auflösung mehr, nur laufend
            ("aus_Verkehrswert", lj(2028), 1_200_000),
            ("aus_Buchwert", lj(2028), 480_000),  # AK G+B 0
            ("aus_StilleReserven", lj(2028), 720_000),
            ("aus_Ruecklage", lj(2028), 0),
            ("aus_LatenteSteuer", lj(2028), 216_000),
            ("aus_VermoegenNetto", lj(2028), 1_225_900),
            ("asb_Verkehrswert", lj(2028), 1_456_560),  # Baseline ohne Neuobjekt
        ]),
        # Verlust Gebäude −136.000 wirkt sofort, Rücklage G+B 920.000 wird 2034 aufgelöst
        ("Etappe 7: Verlustvortrag",
         Modell(objekte=[obj], verkaeufe=[Verkauf("OBJ-001", 2030, preis=1_400_000,
                                                  anteil_gub=0.8, nutzung_6b="ja")],
                parameter=OHNE_ZINS), [
            ("liq_Verkauf", lj(2030), -136_000),
            ("liq_ZvE", lj(2030), -95_884.57),
            ("liq_Bemessung", lj(2030), 0),
            ("liq_Steuer", lj(2030), 0),
            ("liq_Vortrag", lj(2030), 95_884.57),
            ("liq_Vortrag", lj(2033), 95_884.57),
            ("liq_Verkauf", lj(2034), 1_140_800),  # 920.000 + 24 % Zuschlag
            ("liq_VortragGenutzt", lj(2034), 95_884.57),
            ("liq_Bemessung", lj(2034), 1_044_915.43),
            ("liq_Steuer", lj(2034), 313_474.63),
            ("liq_Vortrag", lj(2034), 0),
            ("aus_Vortrag", lj(2030), 95_884.57),
            ("aus_LatenteSteuer", lj(2030), 247_234.63),  # (920.000 − 95.884,57) × 30 %
        ]),
        # Etappe 8: Abnahmefall der Reinvestition, Handrechnung mit eigenem Nachbau.
        # A: Rücklage 720.000 auf das Neuobjekt, AfA-Basis 480.000, AfA 14.400.
        # B: Gewinn 720.000 sofort versteuert, Neuobjekt entfällt, nur Geld.
        # C: Gewinn sofort versteuert, Neuobjekt mit voller AfA-Basis 840.000, AfA 25.200.
        # Ohne Zins gleichen A und C sich aus: § 6b stundet die Steuer nur.
        ("Etappe 8: Abnahme A, B, C ohne Alternativrendite",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")],
                parameter=OHNE_ZINS), [
            ("ne_MitQuelle", 0, 1),
            ("prg_MitQuelle", prg_neu(1, 2029), 1),
            ("prg_AfAOhne6b", prg_neu(1, 2029), 25_200),
            ("prg_BuchwertOhne6b", prg_neu(1, 2028), 840_000),
            ("prg_BuchwertGuBOhne6b", prg_neu(1, 2028), 360_000),
            ("prg_AfAOhne6b", prg(1, 2027), 16_000),
            ("liq_Steuer", lj(2027), 11_100),
            ("lvb_Verkauf", lj(2027), 720_000),
            ("lvb_Steuer", lj(2027), 227_100),  # (37.000 + 720.000) × 30 %
            ("lvb_Kum", lj(2027), 1_225_900),
            ("lvb_Kauf", lj(2028), 0),
            ("lvb_Einnahmen", lj(2029), 0),
            ("lvb_Kum", lj(2046), 1_225_900),
            ("lvc_Steuer", lj(2027), 227_100),
            ("lvc_Kauf", lj(2028), 1_200_000),
            ("lvc_AfA", lj(2029), 25_200),
            ("liq_AfA", lj(2029), 14_400),
            ("avc_Buchwert", lj(2028), 1_200_000),
            ("avc_StilleReserven", lj(2028), 0),
            ("avc_VermoegenNetto", lj(2028), 1_225_900),
            ("aus_VermoegenNetto", lj(2028), 1_225_900),
            ("avb_VermoegenNetto", lj(2028), 1_225_900),
            ("aus_LatenteSteuer", lj(2046), 447_928.65),
            ("avc_LatenteSteuer", lj(2046), 290_248.65),
            ("liq_Kum", lj(2046), 1_044_217),
            ("lvc_Kum", lj(2046), 886_537),
            ("vg_A", 0, 2_310_183.85),
            ("vg_B", 0, 1_225_900),
            ("vg_C", 0, 2_310_183.85),
            ("vg_DiffC", 0, 0),
            ("vg_DiffB", 0, 1_084_283.85),
            ("vgj_C", lj(2046), 2_310_183.85),
            ("vgj_Jahr", lj(2046), 2046),
        ]),
        # Gleiche Eingaben mit 3 % Alternativrendite (Standard): A liegt vor C, der Vorsprung
        # ist der Zins auf die gestundete Steuer. Zins 2028 in A: 1.441.900 × 3 %
        ("Etappe 8: A, B, C mit Alternativrendite 3 %",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")]), [
            ("liq_Zins", lj(2027), 0),
            ("liq_Zins", lj(2028), 43_257),
            ("liq_Kum", lj(2028), 272_179.90),
            ("lvb_Zins", lj(2028), 36_777),
            ("lvc_Kum", lj(2028), 51_643.90),
            ("aus_Zins", lj(2028), 43_257),
            ("liq_Zins", lj(2046), 38_151.36),
            ("lvb_Zins", lj(2046), 53_461.32),
            ("vg_A", 0, 2_615_589.88),
            ("vg_B", 0, 1_819_466.84),
            ("vg_C", 0, 2_522_678.66),
            ("vg_DiffC", 0, 92_911.22),
            ("vg_A", 2, 1_349_623.03),  # Liquidität Ende
            ("vg_C", 3, 290_248.65),    # latente Steuer Ende
        ]),
        # Ohne § 6b-Rücklage sind A, B und C gleich; ein Neuobjekt ohne Quelle bleibt in B.
        # Ohne Verkauf ist A gleich der Baseline, auch mit Zins: 2028 Zins 83.086 × 3 %
        ("Etappe 8: ohne Rücklage sind alle Szenarien gleich",
         Modell(objekte=[obj, obj2]), [
            ("liq_Zins", lj(2028), 2_492.58),
            ("lqb_Zins", lj(2028), 2_492.58),
            ("vg_DiffB", 0, 0),
            ("vg_DiffC", 0, 0),
            ("vg_DiffBaseline", 0, 0),
            ("vg_DiffBaseline", 2, 0),
        ]),
        ("Etappe 8: Neuobjekt ohne Rücklage bleibt in Szenario B",
         Modell(objekte=[obj], neuobjekte=[
             Neuobjekt("NEU-001", 2027, kaufpreis=1_000_000, anteil_gub=0.2, afa_satz=0.03,
                       mietrendite=0.05)]), [
            ("ne_MitQuelle", 0, 0),
            ("lvb_Kauf", lj(2027), 1_000_000),
            ("lvb_AfA", lj(2028), 40_000),  # 16.000 + 24.000
            ("vg_DiffB", 0, 0),
            ("vg_DiffC", 0, 0),
            ("vg_Baseline", 16, 0),  # Kauf Neuobjekte
            ("vg_A", 16, 1_000_000),
        ]),
        # BWA-Ausgabe und Sonderbereich (Projektplan Abschnitt 18), Abnahmefall der
        # Reinvestition mit 3 % Alternativrendite. Das Summenblatt stimmt mit Liquidität
        # und Auswertung überein. Vergleichsjahr 2029, das erste volle Jahr nach dem Kauf:
        # halten Miete 60.000 × 1,02³ = 63.672,48, Erhaltung 8.000 × 1,025³ = 8.615,125,
        # AfA 16.000; Alternative Miete 61.200, Erhaltung 12.300, AfA 14.400, Zins auf die
        # Kapitalanlage 1,4 Mio − 1,2 Mio = 200.000 × 3 % = 6.000
        ("BWA: Abnahmefall Verkauf und Reinvestition",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")]), [
            ("bwa_Jahr", 0, 2027),
            *[(f"bwa_{nr}", lj(j), Wie(name, lj(j)))
              for j in (2027, 2028, 2029, 2046)
              for nr, name in ((1345, "liq_ZvE"), (1355, "liq_Steuer"),
                               (1380, "aus_NachSteuer"), (1240, "liq_AfA"))],
            ("bwa_1323", lj(2027), 720_000),
            ("bwa_1312", lj(2027), 720_000),
            ("bwa_1020", lj(2027), 61_200),
            ("bwa_1020", lj(2029), 61_200),
            ("bwa_1322", lj(2028), 43_257),
            ("bwa_1310", lj(2028), 0),
            ("BWA:OBJ-001:1020", 2026, 60_000),
            ("BWA:OBJ-001:1280", 2026, 8_000),
            # Monatsspalten H–S gruppiert und eingeklappt, Jahres- und Planspalten offen
            *[(f"{EINGEKLAPPT}{blatt}", spalte, soll)
              for blatt in ("OBJ-001", "NEU-001", "Alle Objekte")
              for spalte, soll in ((7, (0, False)), (8, (1, True)), (19, (1, True)),
                                   (20, (0, False)), (21, (0, False)))],
            ("BWA:OBJ-001:1020", 2027, 61_200),
            ("BWA:OBJ-001:1240", 2027, 16_000),
            ("BWA:OBJ-001:1323", 2027, 720_000),
            ("BWA:OBJ-001:1312", 2027, 720_000),
            ("BWA:OBJ-001:1345", 2027, 37_000),
            ("BWA:OBJ-001:1020", 2028, 0),
            ("BWA:NEU-001:1020", 2028, 0),
            ("BWA:NEU-001:1020", 2029, 61_200),
            ("BWA:NEU-001:1250", 2029, 12_300),
            ("BWA:NEU-001:1240", 2029, 14_400),
            ("BWA:NEU-001:1300", 2029, 34_500),
            ("SB1:rechnet mit (1 = ja)", 1, 1),
            ("SB1:Vergleichsjahr: erstes volles Jahr nach Verkauf und Kauf", 1, 2029),
            ("SB1:Bewertung/Erlös (Verkaufspreis)", 1, 1_400_000),
            ("SB1:Reinvestition (Kaufpreis und Nebenkosten der Neuobjekte)", 1, 1_200_000),
            ("SB1:Übertrag § 6b EStG", 1, -720_000),
            ("SB1:Steuer auf den Gewinn ca. (Verkaufsjahr, Auflösung im Fristjahr)", 1, 0),
            (f"SB1:{KAPITALANLAGE}", 1, 200_000),
            ("SB1:Mieten", 1, 63_672.48),
            ("SB1:Mieten", 2, 61_200),
            ("SB1:Zinsertrag Kapitalanlage", 2, 6_000),
            ("SB1:./. Erhaltung (mit Alterung und Großmaßnahme)", 1, -8_615.125),
            ("SB1:./. Erhaltung (mit Alterung und Großmaßnahme)", 2, -12_300),
            ("SB1:./. Abschreibungen (Steuerbilanz)", 1, -16_000),
            ("SB1:./. Abschreibungen (Steuerbilanz)", 2, -14_400),
            ("SB1:= vorläufiges Ergebnis", 1, 39_057.355),
            ("SB1:= vorläufiges Ergebnis", 2, 40_500),
            ("SB1:= Cash Flow", 1, 55_057.355),
            ("SB1:= Cash Flow", 2, 54_900),
            ("SB1:vorläufiges Ergebnis", 3, 1_442.645),
            ("SB1:Steuern ca.", 1, -11_717.2065),
            ("SB1:liquider Überschuss nach Steuern ca.", 2, 42_750),
            ("SB1:Veräußerungspreis", 2, 700_000),
            ("SB1:./. Buchwert (Gebäude zum Ende des Verkaufsjahrs, G+B = AK)", 1, -680_000),
            ("SB1:./. Buchwert (Gebäude zum Ende des Verkaufsjahrs, G+B = AK)", 3, -480_000),
            ("SB1:= Veräußerungsgewinn", 2, 500_000),
            ("SB1:= Veräußerungsgewinn", 3, 220_000),
            ("SB1:Rücklage § 6b (gebildet im Verkaufsjahr)", 1, 720_000),
            ("SB1:übertragen auf Neuobjekte (G+B-Gewinn auf G+B und Gebäude)", 2, 500_000),
            ("SB1:übertragen auf Neuobjekte (G+B-Gewinn auf G+B und Gebäude)", 3, 220_000),
            ("SB1:Fristjahr", 1, 2031),
            ("SB1:aufgelöst im Fristjahr", 1, 0),
            ("SB2:Vorgang 2: kein Verkauf erfasst", 0, "Vorgang 2: kein Verkauf erfasst"),
            ("SB2:rechnet mit (1 = ja)", 1, 0),
            ("SB2:Mieten", 1, None),
            ("KAUF:NeuID (Kostenstelle)", 1, "NEU-001"),
            ("KAUF:NeuID (Kostenstelle)", 2, None),
            ("KAUF:AK G+B", 1, 360_000),
            ("KAUF:ü2 G+B-Gewinn auf G+B", 1, -360_000),
            ("KAUF:ü3 G+B-Gewinn auf Gebäude", 1, -140_000),
            ("KAUF:Übertrag § 6b EStG gesamt", 1, -720_000),
            ("KAUF:AfA-Bemessungsgrundlage Gebäude", 1, 480_000),
            ("KAUF:AfA-Methode", 1, "linear"),
            ("KAUF:erstes volles Jahr", 1, 2029),
            ("KAUF:Mieten", 1, 61_200),
            ("KAUF:./. Abschreibungen (Steuerbilanz)", 1, -14_400),
            ("KAUF:= vorläufiges Ergebnis", 1, 34_500),
            ("KAUF:= Cash Flow", 1, 48_900),
            # Halten liegt nach 20 Jahren vorn: Endvermögen Baseline über A 2.615.589,88
            ("start_Empfehlung", 0, "Halten: kein Verkaufsszenario erreicht ein höheres "
             "Endvermögen."),
            ("start_Vorsprung", 0, 0),
            ("start_Werte", 1, 2_615_589.88),
        ]),
        # Kostenstellen aus der Vorlage: weitere Ausgaben nach dem Anteil der Kostenart im
        # Basisjahr. KSt 1: 1140 1.200, 1150 1.800, 1260 600, je × 1,02; KSt 2: Basis 9.300
        ("BWA: Kostenstellen eingelesen, Kostenarten aufgeteilt", vorlage_eingelesen(), [
            ("BWA:KSt 1:1020", 2025, 58_800),
            ("BWA:KSt 1:1020", 2026, 60_000),
            ("BWA:KSt 1:1150", 2026, 1_800),
            ("BWA:KSt 1:1020", 2027, 61_200),
            ("BWA:KSt 1:1140", 2027, 1_224),
            ("BWA:KSt 1:1150", 2027, 1_836),
            ("BWA:KSt 1:1260", 2027, 612),
            ("BWA:KSt 1:1280", 2027, 27_872),   # mit AfA 16.000 und Erhaltung 8.200
            ("BWA:KSt 2:1120", 2027, 2_448),
            ("BWA:KSt 2:1260", 2027, 1_122),
            ("BWA:Alle Objekte:1020", 2026, 180_000),
            ("BWA:Alle Objekte:1310", 2025, 12_740),  # (4.000 + 9.000) × 0,98
            ("bwa_1120", lj(2027), 2_448),
            ("bwa_1260", lj(2027), 1_734),
            ("bwa_1020", lj(2027), 183_600),
        ]),
        # Kostenstellenblatt führt: im Blatt KSt 1 Basisjahr Miete 66.000 und 1150 2.800
        # getippt. Objekte, Prognose und Alle Objekte folgen; Vorjahr bleibt Ist.
        # weitere Ausgaben 1.200 + 2.800 + 600 = 4.600, 1150 im Plan nach Anteil 2.800 / 4.600
        ("BWA: Kostenstellenblatt führt im Basisjahr",
         vorlage_eingelesen(basis={1020: 66_000, 1150: 2_800}), [
            ("BWA:KSt 1:1020", 2026, 66_000),
            ("BWA:KSt 1:1020", 2025, 58_800),
            ("BWA:KSt 1:1150", 2026, 2_800),
            ("BWA:KSt 1:1280", 2026, 28_600),
            ("BWA:KSt 1:1345", 2026, 33_400),   # − Zins 4.000 lt. Ist
            ("obj_MieteBasis", 0, 66_000),
            ("obj_AusgBasis", 0, 4_600),
            ("obj_ErhBasis", 0, 8_000),
            ("obj_AfABWA", 0, 16_000),
            ("obj_EinnBasis", 1, 1_500),        # KSt 2
            ("BWA:Alle Objekte:1020", 2026, 186_000),
            ("BWA:Alle Objekte:1150", 2026, 6_000),
            ("BWA:Alle Objekte:1310", 2026, 13_000),
            ("BWA:KSt 1:1020", 2027, 67_320),
            ("BWA:KSt 1:1150", 2027, 2_856),
            ("prg_Miete", prg(1, 2027), 67_320),
            ("pr_Anzahl", pr("kst_abweichung"), 0),
        ]),
        # im Blatt Objekte überschrieben: nur die Prognose rechnet damit, das
        # Kostenstellenblatt und Alle Objekte nicht; orange und Warnung
        ("Objekte: Basiswert überschrieben, weicht vom Kostenstellenblatt ab",
         Getippt(vorlage_eingelesen(), {("Objekte", f"{spalte('miete')}2"): 66_000}), [
            ("obj_MieteBasis", 0, 66_000),
            ("prg_Miete", prg(1, 2027), 67_320),
            ("BWA:KSt 1:1020", 2026, 60_000),
            ("BWA:Alle Objekte:1020", 2026, 180_000),
            ("pr_Anzahl", pr("kst_abweichung"), 1),
            ("pr_Ergebnis", pr("kst_abweichung"), WARNUNG),
        ]),
        # Makro schreibt die überschriebenen Werte ins Kostenstellenblatt: Miete 66.000,
        # weitere Ausgaben 4.600 statt 3.600, Differenz 1.000 auf 1260 (600 -> 1.600)
        ("Objekte: Makro überträgt überschriebene Werte ins Kostenstellenblatt",
         Getippt(vorlage_eingelesen(), {("Objekte", f"{spalte('miete')}2"): 66_000,
                                        ("Objekte", f"{spalte('weitere_ausgaben')}2"): 4_600}), [
            ("BWA:KSt 1:1020", 2026, 66_000),
            ("BWA:KSt 1:1260", 2026, 1_600),
            ("BWA:KSt 1:1150", 2026, 1_800),
            ("BWA:KSt 1:1280", 2026, 28_600),
            ("BWA:Alle Objekte:1020", 2026, 186_000),
            ("obj_MieteBasis", 0, 66_000),
            ("obj_AusgBasis", 0, 4_600),
            ("prg_Miete", prg(1, 2027), 67_320),
            ("pr_Anzahl", pr("kst_abweichung"), 0),
        ], [
            ("modObjekte", "InKostenstelleUebernehmen", (), 2),
            ("modObjekte", "InKostenstelleUebernehmen", (), 0),   # nichts mehr abweichend
        ]),
        # Neukauf-Kostenstelle KSt 31 (hinter KSt 9999): Miete 2028 52.000, 2029 53.000 lt.
        # Planspalten, danach × 1,02; Erhaltung 2.000 und 1150 1.000 im Basisjahr, fortge-
        # schrieben mit 2,5 % bzw. 2 %. NEU-2 nennt eine unbekannte Kostenstelle: Mietrendite
        ("Neuobjekte: Prognose aus der Neukauf-Kostenstelle",
         vorlage_eingelesen(
             neukauf=[("KSt 31", "Neubau Nord", {1020: 50_000, 1250: 2_000, 1150: 1_000})],
             plan={"KSt 31": {1020: {2028: 52_000, 2029: 53_000}}},
             neuobjekte=[Neuobjekt("NEU-1", 2027, kaufpreis=1_000_000, anteil_gub=0.2,
                                   afa_satz=0.02, mietrendite=0.04, erhaltungsquote=0.005,
                                   kst="KSt 31"),
                         Neuobjekt("NEU-2", 2027, kaufpreis=500_000, anteil_gub=0.2,
                                   afa_satz=0.02, mietrendite=0.04, erhaltungsquote=0.005,
                                   kst="KSt 99")]), [
            ("nk_ID", 0, "KSt 31"),
            ("obj_ID", 2, "KSt 9999"),
            ("obj_ID", 3, None),                 # KSt 31 ist kein Bestandsobjekt
            ("prg_Miete", prg_neu(1, 2027), 0),
            ("prg_Miete", prg_neu(1, 2028), 52_000),
            ("prg_Miete", prg_neu(1, 2029), 53_000),
            ("prg_Miete", prg_neu(1, 2030), 54_060),
            ("prg_Erhaltung", prg_neu(1, 2028), 2_101.25),
            ("prg_Ausgaben", prg_neu(1, 2028), 1_040.40),
            ("prg_Einnahmen", prg_neu(1, 2028), 0),
            ("prg_AfA", prg_neu(1, 2028), 16_000),   # AK Gebäude 800.000 × 2 %
            ("prg_Miete", prg_neu(2, 2028), 20_400),
            ("BWA:KSt 31:1020", 2026, 50_000),     # eigenes Blatt, Datenbasis und Blatt NEU-1
            ("BWA:KSt 31:1240", 2028, 16_000),     # AfA aus dem Modell
            ("BWA:KSt 31:1310", 2028, 0),          # ohne Kredit keine Zinsen
            ("BWA:NEU-2:1020", 2028, 20_400),      # unbekannte Kostenstelle: eigenes Blatt
            ("BWA:KSt 31:1020", 2028, 52_000),
            ("BWA:KSt 31:1280", 2026, 3_000),      # Summenzeile als Formel
            ("BWA:Alle Objekte:1020", 2026, 180_000),   # Neukauf zählt nicht zum Bestand
            ("pr_Anzahl", pr("neukauf_kst"), 1),
        ]),
        # im Blatt Neukauf-KSt getippt: Miete 2030 60.000, ab 2031 × 1,02; im Blatt KSt 31
        # Basisjahr Miete 55.000 statt 50.000
        ("Neuobjekte: Änderung der Planwerte einer Neukauf-Kostenstelle",
         Getippt(vorlage_eingelesen(
             neukauf=[("KSt 31", "Neubau Nord", {1020: 50_000})],
             neuobjekte=[Neuobjekt("NEU-1", 2027, kaufpreis=1_000_000, anteil_gub=0.2,
                                   afa_satz=0.02, mietrendite=0.04, erhaltungsquote=0.005,
                                   kst="KSt 31")]),
             {("Neukauf-KSt", f"{get_column_letter(nk_spalten()[0] + 4)}{NK_ERSTE}"): 60_000,
              ("KSt 31", f"T{ZEILEN[1020]}"): 55_000}), [
            ("prg_Miete", prg_neu(1, 2029), 58_366.44),   # 55.000 × 1,02³ fortgeschrieben
            ("prg_Miete", prg_neu(1, 2030), 60_000),
            ("prg_Miete", prg_neu(1, 2031), 61_200),
            ("BWA:KSt 31:1020", 2029, 58_366.44),
            ("BWA:KSt 31:1020", 2030, 60_000),
            ("BWA:KSt 31:1020", 2031, 61_200),
        ]),
        # Neukauf-Kostenstelle ohne Buchungen (Miete und Erhaltung 0): Das Neuobjekt rechnet
        # mit Mietrendite 3 % und Erhaltungsquote 0,5 %, das Blatt KSt 31 zeigt die Werte.
        # Miete 2028 1 Mio × 3 % × 1,02 = 30.600; Erhaltung 5.000 × 1,025 = 5.125; weitere
        # Ausgaben 1150 und 1260 je 1.000 im Basisjahr × 1,02² = 1.040,40 je Kostenart.
        # KSt 32 ohne Neuobjekt zeigt ihre Planwerte: 10.000 × 1,02² = 10.404
        ("Neuobjekte: Blatt der Neukauf-Kostenstelle zeigt Miete aus der Mietrendite",
         Getippt(vorlage_eingelesen(
             neukauf=[("KSt 31", "Neubau Nord", {1150: 1_000, 1260: 1_000}),
                      ("KSt 32", "Neubau Süd", {1020: 10_000})],
             neuobjekte=[Neuobjekt("NEU-1", 2027, kaufpreis=1_000_000, anteil_gub=0.2,
                                   afa_satz=0.02, mietrendite=0.03, erhaltungsquote=0.005,
                                   kst="KSt 31")]),
             {("KSt 31", f"T{ZEILEN[1020]}"): 0, ("KSt 31", f"T{ZEILEN[1250]}"): 0}), [
            ("prg_Miete", prg_neu(1, 2028), 30_600),
            ("prg_Erhaltung", prg_neu(1, 2028), 5_125),
            ("prg_Ausgaben", prg_neu(1, 2028), 2_080.80),
            ("BWA:KSt 31:1020", 2027, 0),          # vor dem Kauf keine Miete
            ("BWA:KSt 31:1020", 2028, 30_600),
            ("BWA:KSt 31:1020", 2029, 31_212),
            ("BWA:KSt 31:1250", 2028, 5_125),
            ("BWA:KSt 31:1150", 2028, 1_040.40),
            ("BWA:KSt 31:1260", 2028, 1_040.40),
            ("BWA:KSt 31:1240", 2028, 16_000),
            ("BWA:KSt 32:1020", 2028, 10_404),
            ("zuo_Differenz", lj(2028), 0),
        ]),
        # Finanzierung: AK 2,1 Mio − Nettoerlös 1,4 Mio = Bedarf 700.000, ganz per Kredit.
        # Annuität 4 % + 2 % = 42.000: 2029 Zins 28.000, Tilgung 14.000, Restschuld 686.000;
        # 2030 Zins 27.440, Tilgung 14.560. Übertragung ü1 220.000, ü2 500.000
        ("Finanzierung: Annuitätenkredit über den Finanzierungsbedarf",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=2_000_000, nebenkosten=100_000,
                                      anteil_gub=0.3, afa_satz=0.03, mietrendite=0.05,
                                      erhaltungsquote=0.01, quelle="RL-OBJ-001",
                                      fin_art="Kredit", kredit_zins=0.04, tilgung=0.02)],
                parameter=OHNE_ZINS), [
            ("ne_Status", 0, "OK"),
            ("ne_AKGesamt", 0, 2_100_000),
            ("ne_Erloes", 0, 1_400_000),
            ("ne_Bedarf", 0, 700_000),
            ("ne_Kredit", 0, 700_000),
            ("ne_Eigen", 0, 0),
            ("ne_Ue1", 0, 220_000),
            ("ne_Ue2", 0, 500_000),
            ("ne_Ue3", 0, 0),
            ("ne_AfABasis", 0, 1_250_000),
            ("dl_ID", 0, "NEU-001"),
            ("dl_Rate", 0, 42_000),
            ("dl_Art", 0, "Annuität"),
            ("dl_Getilgt", 0, "nach 2046"),
            ("dl_ID", 1, None),
            ("prg_Restschuld", prg_neu(1, 2028), 700_000),
            ("prg_KreditZins", prg_neu(1, 2028), 0),
            ("prg_KreditZins", prg_neu(1, 2029), 28_000),
            ("prg_Tilgung", prg_neu(1, 2029), 14_000),
            ("prg_Restschuld", prg_neu(1, 2029), 686_000),
            ("prg_KreditZins", prg_neu(1, 2030), 27_440),
            ("prg_Tilgung", prg_neu(1, 2030), 14_560),
            ("prg_Restschuld", prg(1, 2029), 0),
            ("liq_Kauf", lj(2028), 2_100_000),
            ("liq_Kredit", lj(2028), 700_000),
            ("liq_KreditZins", lj(2029), 28_000),
            ("liq_Tilgung", lj(2029), 14_000),
            ("liq_Restschuld", lj(2029), 686_000),
            ("aus_Restschuld", lj(2029), 686_000),
            ("aus_KreditZins", lj(2029), 28_000),
            ("lvb_Kredit", lj(2028), 0),           # B: Neuobjekt mit Rücklage entfällt
            ("lvb_Restschuld", lj(2029), 0),
            ("lvc_Kredit", lj(2028), 700_000),     # C: gleiche Finanzierung
            ("lvc_KreditZins", lj(2029), 28_000),
            ("lqb_Restschuld", lj(2029), 0),
            ("vg_A", 17, 340_964.22),          # Restschuld Ende 2046
            ("vg_Baseline", 17, 0),
            ("bwa_1310", lj(2029), 28_000),
            ("BWA:NEU-001:1310", 2029, 28_000),
            ("BWA:NEU-001:1310", 2028, 0),
            ("zuo_Differenz", lj(2029), 0),
            ("zuo_Differenz", lj(2030), 0),
            ("bwah_neu_zins", lj(2029), -28_000),
            ("KAUF:= Finanzierungsbedarf", 1, 700_000),
            ("KAUF:davon Kredit", 1, 700_000),
            ("KAUF:./. Zinsen Kredit", 1, -28_000),
            ("KAUF:./. Tilgung Kredit", 1, -14_000),
            ("SB1:Kredit der Neuobjekte", 1, 700_000),
            (f"SB1:{KAPITALANLAGE}", 1, 0),
            ("SB1:./. Zinsen Darlehen", 2, -28_000),
            ("SB1:./. Tilgungen", 2, -14_000),
            ("pr_Anzahl", pr("restschuld"), 1),
            ("pr_Anzahl", pr("neuobjekte"), 0),
        ]),
        # Zwei Verkäufe mit je Rücklage Gebäude 220.000, G+B 500.000 und Nettoerlös 1,4 Mio.
        # NEU-001 (AK Gebäude 2,4 Mio, G+B 600.000) nimmt beide: ü1 220.000 + 220.000,
        # ü2 500.000 + 100.000, ü3 aus Quelle 2 400.000; Erlös 2,8 Mio, Bedarf 200.000,
        # Kredit 150.000 linear 25 % zu 5 %, getilgt 2032. NEU-002 endfällig nach 5 Jahren.
        ("Finanzierung: zwei Quellen, linearer und endfälliger Kredit, Statusfälle",
         Modell(objekte=[abnahme4, dataclasses.replace(abnahme4, objekt_id="OBJ-002")],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja"),
                           Verkauf("OBJ-002", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[
                    Neuobjekt("NEU-001", 2028, kaufpreis=3_000_000, anteil_gub=0.2,
                              afa_satz=0.03, quelle="RL-OBJ-001", quelle2="RL-OBJ-002",
                              fin_art="Kredit", kredit_betrag=150_000, kredit_zins=0.05,
                              tilgungsart="linear", tilgung=0.25),
                    Neuobjekt("NEU-002", 2029, kaufpreis=500_000, anteil_gub=0.2, afa_satz=0.02,
                              fin_art="Kredit", kredit_zins=0.03, tilgungsart="endfällig",
                              laufzeit=5),
                    Neuobjekt("NEU-003", 2030, kaufpreis=100_000, anteil_gub=0.2, afa_satz=0.02,
                              fin_art="Kredit"),
                    Neuobjekt("NEU-004", 2030, kaufpreis=100_000, anteil_gub=0.2, afa_satz=0.02,
                              kredit_zins=0.05),
                    Neuobjekt("NEU-005", 2030, kaufpreis=100_000, anteil_gub=0.2, afa_satz=0.02,
                              quelle="RL-OBJ-001", quelle2="RL-OBJ-001"),
                    Neuobjekt("NEU-006", 2031, kaufpreis=100_000, anteil_gub=0.2, afa_satz=0.02,
                              fin_art="Kredit", kredit_zins=0.04, tilgungsart="endfällig")],
                parameter=OHNE_ZINS), [
            ("ne_Status", 0, "OK"),
            ("ne_Q1Ue1", 0, 220_000),
            ("ne_Q2Ue1", 0, 220_000),
            ("ne_Q1Ue2", 0, 500_000),
            ("ne_Q2Ue2", 0, 100_000),
            ("ne_Q1Ue3", 0, 0),
            ("ne_Q2Ue3", 0, 400_000),
            ("ne_Ue1", 0, 440_000),
            ("ne_Ue2", 0, 600_000),
            ("ne_Ue3", 0, 400_000),
            ("ne_UeGesamt", 0, 1_440_000),
            ("ne_AfABasis", 0, 1_560_000),
            ("ne_AKGuB", 0, 0),
            ("ne_MitQuelle", 0, 1),
            ("rl_UebGeb", 1, 220_000),
            ("rl_UebGuB", 1, 500_000),
            ("rl_Aufloesung", 0, 0),
            ("rl_Aufloesung", 1, 0),
            ("rls_Uebertragung", rls(2028), 1_440_000),
            ("ne_Q1Erloes", 0, 1_400_000),
            ("ne_Q2Erloes", 0, 1_400_000),
            ("ne_Erloes", 0, 2_800_000),
            ("ne_Bedarf", 0, 200_000),
            ("ne_Kredit", 0, 150_000),
            ("ne_Eigen", 0, 50_000),
            ("dl_Rate", 0, 37_500),
            ("prg_KreditZins", prg_neu(1, 2029), 7_500),
            ("prg_Restschuld", prg_neu(1, 2029), 112_500),
            ("prg_KreditZins", prg_neu(1, 2032), 1_875),
            ("prg_Tilgung", prg_neu(1, 2032), 37_500),
            ("prg_Restschuld", prg_neu(1, 2032), 0),
            ("prg_KreditZins", prg_neu(1, 2033), 0),
            ("prg_Tilgung", prg_neu(1, 2033), 0),
            ("dl_Getilgt", 0, 2032),
            ("ne_Kredit", 1, 500_000),              # Betrag leer: ganzer Bedarf
            ("prg_KreditZins", prg_neu(2, 2030), 15_000),
            ("prg_Tilgung", prg_neu(2, 2033), 0),
            ("prg_Tilgung", prg_neu(2, 2034), 500_000),
            ("prg_Restschuld", prg_neu(2, 2034), 0),
            ("prg_KreditZins", prg_neu(2, 2035), 0),
            ("dl_Getilgt", 1, 2034),
            ("ne_Status", 2, "Kredit: Zinssatz fehlt"),
            ("ne_Status", 3, "Kreditangaben ohne Finanzierung Rest = Kredit, kein Kredit"),
            ("ne_Kredit", 3, 0),
            ("ne_Eigen", 3, 100_000),
            ("dl_ID", 3, None),
            ("ne_Status", 4, "Quelle 2: doppelt, keine Übertragung"),
            ("ne_UeGesamt", 4, 0),
            ("ne_Erloes", 4, 0),
            ("ne_Status", 5, "Kredit: Laufzeit fehlt (endfällig)"),
            ("dl_Getilgt", 5, "nach 2046"),
            ("liq_Kredit", lj(2028), 150_000),
            ("liq_KreditZins", lj(2030), 20_625),
            ("liq_KreditZins", lj(2032), 20_875),
            ("liq_Tilgung", lj(2034), 500_000),
            ("zuo_Differenz", lj(2032), 0),
            ("pr_Anzahl", pr("neuobjekte"), 4),
            ("pr_Anzahl", pr("restschuld"), 2),
            ("pr_Anzahl", pr("teiluebertrag"), 0),
            ("KAUF:Quelle 2 RücklageID", 1, "RL-OBJ-002"),
            ("KAUF:davon Eigenmittel", 1, 50_000),
        ]),
        # Darlehen der Bestandsobjekte (Projektplan Abschnitt 28). OBJ-001: Restschuld 500.000,
        # 4 %, Rate 40.000, Verkauf Ende 2029: 2027 Zins 20.000, Tilgung 20.000; 2028 Zins
        # 19.200; 2029 Zins 18.368, Ablösung 459.200, bei Halten Tilgung 21.632, Rest
        # 437.568. OBJ-002: nur Zinsaufwand 10.000, × 0,97 je Jahr. OBJ-003: Restschuld
        # 300.000, Zinsaufwand 12.000: Satz 4 % und Rate 300.000 × 6 % = 18.000 als Annahme.
        # OBJ-004: Zinsaufwand 5.000 ohne Restschuld, Verkauf Ende 2030: Ablösung fehlt
        ("Bestandsdarlehen: Tilgungsplan, Fortschreibung und Ablösung beim Verkauf",
         Modell(objekte=[dataclasses.replace(obj, restschuld=500_000, zinssatz=0.04,
                                             rate=40_000),
                         dataclasses.replace(obj2, zinsen=10_000),
                         dataclasses.replace(obj, objekt_id="OBJ-003", zinsen=12_000,
                                             restschuld=300_000),
                         dataclasses.replace(obj, objekt_id="OBJ-004", zinsen=5_000)],
                verkaeufe=[Verkauf("OBJ-001", 2029, preis=1_400_000),
                           Verkauf("OBJ-004", 2030, preis=1_400_000)],
                parameter=OHNE_ZINS), [
            ("prg_KreditZins", prg(1, 2027), 20_000),
            ("prg_Tilgung", prg(1, 2027), 20_000),
            ("prg_Restschuld", prg(1, 2027), 480_000),
            ("prg_KreditZins", prg(1, 2028), 19_200),
            ("prg_KreditZins", prg(1, 2029), 18_368),
            ("prg_Tilgung", prg(1, 2029), 459_200),          # Ablösung zum Verkauf
            ("prg_Restschuld", prg(1, 2029), 0),
            ("prg_KreditZins", prg(1, 2030), 0),
            ("prg_Tilgung", prg(1, 2030), 0),
            ("prg_TilgungHalten", prg(1, 2029), 21_632),
            ("prg_RestschuldHalten", prg(1, 2029), 437_568),
            ("prg_ZinsHalten", prg(1, 2030), 17_502.72),
            ("prg_KreditZins", prg(2, 2027), 9_700),
            ("prg_KreditZins", prg(2, 2028), 9_409),
            ("prg_Tilgung", prg(2, 2028), 0),
            ("prg_Restschuld", prg(2, 2028), 0),
            ("obj_Zinssatz", 1, 0),                         # ohne Restschuld 0, keine Annahme
            ("obj_Rate", 1, 0),
            ("obj_Annahmen", 1, 0),
            ("obj_Zinssatz", 2, 0.04),
            ("obj_Rate", 2, 18_000),
            ("obj_Annahmen", 2, 2),
            ("prg_Tilgung", prg(3, 2027), 6_000),
            ("prg_Restschuld", prg(3, 2029), 281_270.4),
            ("prg_KreditZins", prg(4, 2030), 4_426.46405),  # 5.000 × 0,97⁴, Verkaufsjahr
            ("prg_KreditZins", prg(4, 2031), 0),
            ("prg_ZinsHalten", prg(4, 2031), 4_293.6701285),
            ("liq_KreditZins", lj(2027), 46_550),
            ("liq_Tilgung", lj(2027), 26_000),
            ("liq_Restschuld", lj(2027), 774_000),
            ("liq_Tilgung", lj(2029), 465_689.6),
            ("liq_KreditZins", lj(2030), 24_530.20815),
            ("lqb_Tilgung", lj(2029), 28_121.6),
            ("lqb_KreditZins", lj(2030), 42_032.92815),
            ("lvb_KreditZins", lj(2027), 46_550),
            ("aus_Restschuld", lj(2029), 281_270.4),
            ("asb_Restschuld", lj(2029), 718_838.4),
            ("bwa_1310", lj(2027), 46_550),
            *[(f"bwa_{nr}", lj(j), Wie(name, lj(j)))
              for j in (2027, 2029, 2030) for nr, name in ((1345, "liq_ZvE"),)],
            ("BWA:OBJ-001:1310", 2027, 20_000),
            ("BWA:OBJ-002:1310", 2026, 10_000),
            ("BWA:OBJ-002:1310", 2027, 9_700),
            ("SB1:Ablösung Darlehen des Objekts (Restschuld Ende Verkaufsjahr)", 1, -437_568),
            # Nettoerlös 1,4 Mio − Steuer 30 % auf 768.000 − Ablösung
            (f"SB1:{KAPITALANLAGE}", 1, 732_032),
            ("SB1:./. Zinsen Darlehen", 1, -17_502.72),
            ("SB1:./. Zinsen Darlehen", 2, 0),
            ("SB1:./. Tilgungen", 1, -22_497.28),
            ("pr_Anzahl", pr("zins_verkauf"), 1),
            ("pr_Anzahl", pr("zins_grob"), 2),
        ]),
        # Neuobjekt mit Neukauf-Kostenstelle: deren Blatt zeigt AfA und Kreditzinsen,
        # Kredit 400.000 zu 5 %: Zins 2028 20.000
        ("Finanzierung: Kreditzinsen im Blatt der Neukauf-Kostenstelle",
         vorlage_eingelesen(
             neukauf=[("KSt 31", "Neubau Nord", {1020: 50_000})],
             neuobjekte=[Neuobjekt("NEU-1", 2027, kaufpreis=1_000_000, anteil_gub=0.2,
                                   afa_satz=0.02, kst="KSt 31", fin_art="Kredit",
                                   kredit_betrag=400_000, kredit_zins=0.05, tilgung=0.03)]), [
            ("ne_Status", 0, "OK"),
            ("ne_Bedarf", 0, 1_000_000),
            ("ne_Eigen", 0, 600_000),
            ("BWA:KSt 31:1310", 2028, 20_000),
            ("BWA:KSt 31:1240", 2028, 16_000),
            ("BWA:KSt 31:1320", 2028, 20_000),
            ("zuo_Differenz", lj(2028), 0),
        ]),
        # AfA-Plan aus BWA 1240 der Kostenstellen-Datei: 2027 15.000, 2028 14.000, 2030 0;
        # 2029 und ab 2031 schreibt das Modell fort (AfA lt. Buchhaltung 16.000)
        ("Selbstdokumentation: Pflichtfelder, Tooltips, eingeklappte Parameter", [testobjekt()], [
            (f"{KOPF}Verkäufe:1", 1, "ObjektID * [Tooltip]"),
            (f"{KOPF}Verkäufe:1", nr(VERKAUF_FELDER, "kosten"), "Verkaufskosten [Tooltip]"),
            (f"{KOPF}Verkäufe:1", nr(VERKAUF_SPALTEN, "nettoerloes", len(VERKAUF_FELDER)),
             "Nettoerlös [Tooltip]"),
            (f"{KOPF}Neuobjekte:1", nr(NEU_FELDER, "kaufpreis"), "Kaufpreis * [Tooltip]"),
            (f"{KOPF}Neuobjekte:1", nr(NEU_FELDER, "kst"), "Kostenstelle Neukauf [Tooltip]"),
            (f"{KOPF}Neuobjekte:1", nr(NEU_SPALTEN, "bedarf", len(NEU_FELDER)),
             "Finanzierungsbedarf [Tooltip]"),
            (f"{KOPF}Objekte:1", 1, "ObjektID * [Tooltip]"),
            (f"{KOPF}Liquidität:2", nr(LIQ_SPALTEN, "zve"),
             "Ergebnis vor Verlustvortrag [Tooltip]"),
            (f"{KOPF}Liquidität:2", 1, "Jahr"),
            (f"{KOPF}Rücklagen:1", 1, "RücklageID [Tooltip]"),
            (f"{KOPF}Parameter:5", 1, "★ Basisjahr (Ist) [Tooltip]"),
            (f"{KOPF}Parameter:13", 1, "Steigerung weitere Ausgaben p. a. [Tooltip]"),
            # § 6b-Abschnitt eingeklappt, Annahmen bei fehlenden Daten (mit Kernparametern) offen
            (f"{ZEILE_EINGEKLAPPT}Parameter", 21, (0, False)),
            (f"{ZEILE_EINGEKLAPPT}Parameter", 22, (1, True)),
            (f"{ZEILE_EINGEKLAPPT}Parameter", 27, (1, True)),
            (f"{ZEILE_EINGEKLAPPT}Parameter", 30, (1, False)),
            (f"{ZEILE_EINGEKLAPPT}Parameter", 38, (1, True)),
            ("par_6bFrist", 0, 4),
            ("par_6bFristNeubau", 0, 6),
            # Auswahllisten: je Zelle genau eine Überprüfung, sonst verwirft Excel sie
            (UEBERLAPPT, 0, []),
            (f"{AUSWAHL}Verkäufe", "A2", ("obj_ID", 1)),
            (f"{AUSWAHL}Verkäufe", "F2", ('"ja,nein"', 1)),
            (f"{AUSWAHL}Neuobjekte", "A2", (None, 1)),
            (f"{AUSWAHL}Neuobjekte", f"{get_column_letter(nr(NEU_FELDER, 'quelle'))}2",
             ("rl_ID", 1)),
            (f"{AUSWAHL}Neuobjekte", f"{get_column_letter(nr(NEU_FELDER, 'kst'))}2",
             ("nk_Liste", 1)),
        ]),
        ("AfA-Plan: geplante Jahre ersetzen die Fortschreibung",
         vorlage_eingelesen(plan={"KSt 1": {1240: {2027: 15_000, 2028: 14_000, 2030: 0}}}), [
            ("afp_ID", 0, "KSt 1"),
            ("afp_Anzahl", 0, 3),
            ("afp_Anzahl", 1, 0),
            ("prg_AfA", prg(1, 2027), 15_000),
            ("prg_AfA", prg(1, 2028), 14_000),
            ("prg_AfA", prg(1, 2029), 16_000),
            ("prg_AfA", prg(1, 2030), 0),
            ("prg_AfA", prg(1, 2031), 16_000),
            ("prg_BuchwertHalten", prg(1, 2030), 480_000 - 45_000),
            ("BWA:KSt 1:1240", 2027, 15_000),
            ("BWA:KSt 1:1240", 2030, 0),
            ("pr_Anzahl", pr("afa_plan"), 1),
        ]),
        # AfA lt. Buchhaltung läuft in der Prognose weiter: 12.000 statt AK × Satz 16.000,
        # bis der Restbuchwert 30.000 verbraucht ist; 0 lt. Buchhaltung = abgeschrieben
        ("AfA: Buchhaltung steuert die Prognose",
         Modell(objekte=[dataclasses.replace(obj, objekt_id="B1", afa_bwa=12_000,
                                             restbuchwert=30_000),
                         Objekt("B2", miete=60_000, afa_bwa=0)]), [
            ("obj_AfAJahr", 0, 12_000),
            ("prg_AfA", prg(1, 2027), 12_000),
            ("prg_AfA", prg(1, 2028), 12_000),
            ("prg_AfA", prg(1, 2029), 6_000),
            ("prg_AfA", prg(1, 2030), 0),
            ("prg_BuchwertHalten", prg(1, 2028), 6_000),
            ("obj_AfAJahr", 1, 0),
            ("obj_Restbuchwert", 1, 0),
            ("prg_AfA", prg(2, 2027), 0),
            ("BWA:B1:1240", 2027, 12_000),
        ]),
        # Zuordnung brutto, Verkauf und Rücklage außerordentlich, Neuobjekt-Miete auf 1090:
        # Preis 1,4 Mio als Ertrag 1351; Buchwertabgang 480.000 + 200.000 und Rücklage
        # 720.000 als Aufwand 1352. Das Ergebnis (1353) bleibt gleich der Liquidität.
        ("BWA: Zuordnung brutto und eigene Zeilen",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3,
                                      afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01,
                                      quelle="RL-OBJ-001")],
                bwa_zuordnung={"verkauf": "brutto", "erloes": 1351, "abgang": 1352,
                               "bildung": 1352, "neu_miete": 1090}), [
            *[(f"bwa_{nr}", lj(j), Wie(name, lj(j)))
              for j in (2027, 2028, 2029)
              for nr, name in ((1353, "liq_ZvE"), (1240, "liq_AfA"))],
            ("bwa_1351", lj(2027), 1_400_000),
            ("bwa_1352", lj(2027), 1_400_000),
            ("bwa_1323", lj(2027), 0),
            ("bwa_1312", lj(2027), 0),
            ("bwa_1020", lj(2029), 0),
            ("bwa_1090", lj(2029), 61_200),
            ("BWA:OBJ-001:1351", 2027, 1_400_000),
            ("BWA:OBJ-001:1352", 2027, 1_400_000),
            ("BWA:NEU-001:1090", 2029, 61_200),
            ("BWA:NEU-001:1020", 2029, 0),
            ("bwah_erloes", lj(2027), 1_400_000),
            ("bwah_gewinn", lj(2027), 0),
            ("zuo_Differenz", 0, 0),
            ("zuo_Differenz", 2, 0),
            ("pr_Anzahl", pr("bwa_zuordnung"), 0),
        ]),
        # ungültige Nr.: der Gewinn 720.000 fehlt in der BWA, die Kontrolle schlägt 2027 an
        ("BWA: Zuordnung mit ungültiger Nr.",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                bwa_zuordnung={"gewinn": 1100}), [
            ("zuo_Differenz", 0, -720_000),
            ("zuo_Differenz", 1, 0),
            ("pr_Anzahl", pr("bwa_zuordnung"), 2),
            ("pr_Ergebnis", pr("bwa_zuordnung"), "Warnung"),
        ]),
        # Auffülllogik: Objekt nur mit Miete. Verkehrswert 60.000 × 20, Gebäudeanteil 75 %,
        # Kauf 2011, AK Gebäude = 900.000 / 1,02^15, ohne AfA lt. Buchhaltung; mit AfA 16.000
        # und 2 % sind es 800.000. Restbuchwert = AK × (1 − 16 × 2 %)
        ("Annahmen: Objekt nur mit Miete und AfA lt. Buchhaltung",
         Modell(objekte=[Objekt("A1", miete=60_000), Objekt("A2", miete=60_000, afa_bwa=16_000)]),
         befund(annahmen=2) + [
            ("obj_Verkehrswert", 0, 1_200_000),
            ("obj_VKQuoteGeb", 0, 0.75),
            ("obj_AfASatz", 0, 0.02),
            ("obj_Kaufjahr", 0, 2011),
            ("obj_AKGebaeude", 0, 668_713.26),
            ("obj_AKGuB", 0, 222_904.42),
            ("obj_Restbuchwert", 0, 454_725.01),
            ("obj_ErhBasis", 0, 6_000),
            ("obj_AKGebaeude", 1, 800_000),
            ("obj_AKGuB", 1, 266_666.67),
            ("obj_Restbuchwert", 1, 544_000),
            ("obj_Status", 0, "OK"),
            ("obj_Annahmen", 0, 11),   # 8 Stammdaten, Baujahr, Großmaßnahme Jahr und Betrag
            ("obj_Kritisch", 1, 0),
            ("prg_AfA", prg(2, 2027), 16_000),
            ("ueb_MitAnnahmen", 0, 2),
            ("start_Empfehlung", 0, "Noch kein Verkauf erfasst: im Blatt Verkäufe Objekt und "
             "Verkaufsjahr eintragen, dann zeigt dieses Blatt die beste Option."),
            ("start_Belastbarkeit", 0, "Belastbar im Rahmen der zentralen Annahmen."),
            ("start_Kritisch", 0, 0),
        ]),
        # Preis = 1,2 Mio × 1,02²; Gewinn Gebäude 936.360 − 512.000, G+B 312.120 − 266.666,67;
        # Neuobjekt 2029 für 1.248.480 / 1,07 plus 7 % Nebenkosten, Quelle die Rücklage
        ("Annahmen: Verkauf ohne Preis, Reinvestition aus dem Verkauf",
         Modell(objekte=[Objekt("A1", miete=60_000), Objekt("A2", miete=60_000, afa_bwa=16_000)],
                verkaeufe=[Verkauf("A2", 2028, nutzung_6b="ja", reinvest="ja")]),
         befund(annahmen=2, kritisch=2) + [
            ("vk_Preis", 0, 1_248_480),
            ("vk_PreisAnnahme", 0, 1),
            ("vk_Gewinn", 0, 469_813.33),
            ("obj_Kritisch", 1, 6),
            ("ne_ID", 0, "NEU-A2"),
            ("ne_Kaufjahr", 0, 2029),
            ("ne_Kaufpreis", 0, 1_166_803.74),
            ("ne_Nebenkosten", 0, 81_676.26),
            ("ne_Quelle", 0, "RL-A2"),
            ("ne_Status", 0, "OK"),
            ("ne_AfABasis", 0, 512_000),
            ("ne_ID", 1, None),
            ("rls_Bestand", rls(2029), 0),       # voll übertragen
            ("start_Belastbarkeit", 0, "Vorläufig: 2 kritische Annahme(n) bei Verkäufen "
             "(orange) durch echte Werte ersetzen."),
            ("start_Werte", 0, Wie("vg_Baseline")),
            ("start_Werte", 1, Wie("vg_A")),
            ("start_Werte", 2, Wie("vg_C")),
            ("start_Werte", 3, Wie("vg_B")),
        ]),
        ("Schnellcheck: § 6b und Reinvestition als Annahme",
         Modell(objekte=[Objekt("A2", miete=60_000, afa_bwa=16_000)],
                verkaeufe=[Verkauf("A2", 2028)], schnellcheck=True), [
            ("vk_6b", 0, "ja"),
            ("vk_Reinvest", 0, "ja"),
            ("vk_6b", 1, None),
            ("rl_ID", 0, "RL-A2"),
            ("ne_ID", 0, "NEU-A2"),
            ("ne_Quelle", 0, "RL-A2"),
            ("ne_AfAMethode", 0, "linear"),
            ("ne_Mietrendite", 0, 0.045),
            ("prg_Miete", prg_neu(1, 2030), 1_166_803.74 * 0.045 * 1.02),
        ]),
        # Erhaltung nach Gebäudealter: Baujahr 2000, Alterung ab 30 Jahren (2030) mit 1,5 %
        # zusätzlich; Großmaßnahme erst mit 50 Jahren, also nach dem Raster
        ("Erhaltung: Alterung ab 30 Jahren",
         Modell(objekte=[dataclasses.replace(obj, baujahr=2000)]), [
            ("obj_SanJahr", 0, 0),
            ("obj_SanBetrag", 0, 0),
            ("prg_Erhaltung", prg(1, 2030), 8_830.50),     # 8.000 × 1,025⁴
            ("prg_Erhaltung", prg(1, 2031), 9_187.03),     # × 1,025 × 1,015
            ("prg_Erhaltung", prg(1, 2046), 16_635.04),    # 8.000 × 1,025²⁰ × 1,015¹⁶
            ("lqb_Ausgaben", lj(2031), 9_187.03),
            ("liq_Ausgaben", lj(2031), 9_187.03),
        ]),
        # Baujahr 1980: im Basisjahr 46 Jahre, Großmaßnahme mit 50 im Jahr 2030 über
        # 1,4 Mio × 50 % × 15 % = 105.000; Verkauf 2028 erspart sie dem Plan, nicht dem Halten
        ("Erhaltung: Großmaßnahme bei Halten, nicht nach Verkauf",
         Modell(objekte=[dataclasses.replace(obj, baujahr=1980, san_jahr=None,
                                             san_betrag=None)],
                verkaeufe=[Verkauf("OBJ-001", 2028, preis=1_400_000)]), [
            ("obj_SanJahr", 0, 2030),
            ("obj_SanBetrag", 0, 105_000),
            ("obj_Annahmen", 0, 2),
            ("prg_ErhaltungHalten", prg(1, 2029), 9_008.65),  # 8.000 × 1,015³ × 1,025³
            ("prg_ErhaltungHalten", prg(1, 2030), 125_272.73),
            ("prg_Erhaltung", prg(1, 2030), 0),
            ("lqb_Ausgaben", lj(2030), 125_272.73),
            ("liq_Ausgaben", lj(2030), 0),
        ]),
        # Baujahr 1975 ist im Basisjahr schon über 50: fällig nach 2 Jahren Vorlauf, also 2029
        ("Erhaltung: überfällige Großmaßnahme nach Vorlauf",
         Modell(objekte=[dataclasses.replace(obj, baujahr=1975, san_jahr=None,
                                             san_betrag=None)]), [
            ("obj_SanJahr", 0, 2029),
            ("obj_SanBetrag", 0, 105_000),
        ]),
        ("Erhaltung: keine Großmaßnahme mit Jahr 0",
         Modell(objekte=[dataclasses.replace(obj, baujahr=1980, san_jahr=0)]), [
            ("obj_SanBetrag", 0, 0),
            ("prg_Erhaltung", prg(1, 2030), 8_000 * 1.015 ** 4 * 1.025 ** 4),
        ]),
        # Neuobjekt: Erhaltung 1 % von 1,2 Mio, in den ersten 10 Jahren nach dem Kauf zur Hälfte
        ("Erhaltung: Neuobjekt mit Anlaufjahren",
         Modell(objekte=[obj], neuobjekte=[
             Neuobjekt("NEU-001", 2028, kaufpreis=1_200_000, anteil_gub=0.3, afa_satz=0.03,
                       mietrendite=0.05, erhaltungsquote=0.01)]), [
            ("prg_Erhaltung", prg_neu(1, 2029), 6_150),
            ("prg_Erhaltung", prg_neu(1, 2038), 7_680.51),
            ("prg_Erhaltung", prg_neu(1, 2039), 15_745.04),
            ("prg_ErhaltungHalten", prg_neu(1, 2039), 0),
        ]),
        # Etappe 9: Plausibilitätsprüfungen melden jeden eingebauten Fehler
        ("Etappe 9: Testobjekt ohne Befund", [obj], befund() + [
            ("par_StatusPruefung", 0, "OK"),
            ("ueb_Pruefung", 0, "OK"),
        ]),
        ("Etappe 9: Fehler in allen Eingabeblättern",
         Modell(objekte=[ohne_miete, obj2],
                verkaeufe=[Verkauf("OBJ-999", 2030, preis=1_000_000)],
                neuobjekte=[Neuobjekt("NEU-001", 2050, kaufpreis=500_000, anteil_gub=0.2,
                                      afa_satz=0.03)],
                parameter={"par_Steuerwelt": "Privat"}),
         befund(objekte=1, verkaeufe=1, neuobjekte=1, steuerwelt=1) + [
            ("par_StatusPruefung", 0, "4 Fehler"),
            ("ueb_Pruefung", 0, "4 Fehler"),
        ]),
        ("Etappe 9: Vorbesitzzeit zu kurz, Rücklage ohne Neuobjekt",
         Modell(objekte=[abnahme4, jung],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja"),
                           Verkauf("OBJ-004", 2027, preis=1_400_000, nutzung_6b="ja")]),
         befund(vorbesitz=1, ohne_reinvest=1)),
        # Gebäudeanteil 210.000 nimmt den Gebäudegewinn 220.000 nicht voll auf
        ("Etappe 9: Rücklage nur teilweise übertragen",
         Modell(objekte=[abnahme4],
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")],
                neuobjekte=[Neuobjekt("NEU-001", 2028, kaufpreis=300_000, anteil_gub=0.3,
                                      afa_satz=0.03, quelle="RL-OBJ-001")]),
         befund(teiluebertrag=1)),
        # Verkäufe 2027 bis 2030: im Zeitraum bis 2030 vier, über der Grenze von drei
        ("Etappe 9: Drei-Objekt-Grenze überschritten",
         Modell(objekte=[obj, obj2, dataclasses.replace(obj, objekt_id="OBJ-003"),
                         dataclasses.replace(obj, objekt_id="OBJ-004")],
                verkaeufe=[Verkauf(f"OBJ-00{i}", 2026 + i, preis=1_400_000)
                           for i in range(1, 5)]),
         befund(grundstueckshandel=1)),
        ("Etappe 9: Drei-Objekt-Grenze mit Zeitraum 3 Jahre eingehalten",
         Modell(objekte=[obj, obj2, dataclasses.replace(obj, objekt_id="OBJ-003"),
                         dataclasses.replace(obj, objekt_id="OBJ-004")],
                verkaeufe=[Verkauf(f"OBJ-00{i}", 2026 + i, preis=1_400_000)
                           for i in range(1, 5)],
                parameter={"par_DOJahre": 3}),
         befund()),
        # Fristjahr 2049 liegt nach dem Raster; die Rücklage hat auch kein Neuobjekt
        ("Etappe 9: Frist endet nach Prognoseende",
         Modell(objekte=[obj],
                verkaeufe=[Verkauf("OBJ-001", 2045, preis=1_400_000, nutzung_6b="ja")]),
         befund(frist_ende=1, ohne_reinvest=1)),
        ("Etappe 9: Kauf ohne Erlös macht Liquidität negativ",
         Modell(objekte=[obj], neuobjekte=[
             Neuobjekt("NEU-001", 2027, kaufpreis=5_000_000, anteil_gub=0.2, afa_satz=0.03)]), [
            ("pr_Ergebnis", pr("liquiditaet"), HINWEIS),
            ("pr_Gesamt", 0, "OK"),  # Hinweise zählen nicht
        ]),
        ("Etappe 9: ohne Verkehrswert nur Hinweis", [obj3], befund(annahmen=1)),
        # Makrofälle: vierter Eintrag sind die Aufrufe der Reihe nach,
        # (Modul, Prozedur, Argumente, erwarteter Rückgabewert; FEHLT = ohne Rückgabewert)
        ("Etappe 9: Makros Objekt anlegen, duplizieren, entfernen", Modell(objekte=[obj, obj2]), [
            ("obj_ID", 0, "OBJ-001"),
            ("obj_ID", 1, "OBJ-005"),            # frei gewordene Zeile wird wieder belegt
            ("obj_ID", 2, "OBJ-003"),
            ("obj_ID", 3, "OBJ-004"),
            ("obj_ID", 4, None),                 # weder OBJ-001 doppelt noch OBJ-006
            ("obj_Name", 2, "Testobjekt"),
            ("obj_Restbuchwert", 2, 480_000),
            ("obj_Status", 2, "OK"),
            ("obj_Status", 1, "Pflichtfeld fehlt"),
            ("obj_Status", 3, "Pflichtfeld fehlt"),
            ("obj_EinnBasis", 1, None),          # Eingaben von OBJ-002 geleert
            ("obj_Kaufjahr", 1, 2011),           # Annahme nach dem Entfernen wiederhergestellt
            ("obj_Kaufjahr", 3, 2011),
            ("prg_ID", prg(3, 2027), "OBJ-003"),  # Kopie rechnet im eigenen Block
            ("prg_Miete", prg(3, 2027), 61_200),
            ("prg_Buchwert", prg(3, 2046), 160_000),
            ("ueb_Plan", ueb(2027), 2 * 1_428_000),
            ("pr_Anzahl", pr("objekte"), 2),
            ("pr_Gesamt", 0, "1 Fehler"),
            (AUSGEBLENDET, 2 + 0 * 20, False),   # Block OBJ-001
            (AUSGEBLENDET, 2 + 2 * 20, False),   # Block OBJ-003
            (AUSGEBLENDET, 2 + 4 * 20, True),    # leerer Block
            (AUSGEBLENDET, 2 + 5 * 20 - 1, True),
            (AUSGEBLENDET, 2 + TEST_KAPAZITAET.objekte * 20, True),  # leerer Neuobjektblock
        ], [
            ("modObjekte", "ObjektDuplizieren", ("OBJ-001", "OBJ-003"), 3),
            ("modObjekte", "ObjektAnlegen", ("OBJ-004",), 4),
            # Err.Raise: In LibreOffice bricht die Funktion ab und liefert 0, nichts wird
            # geschrieben (geprüft über obj_ID); Excel zeigt die Meldung
            ("modObjekte", "ObjektAnlegen", ("OBJ-001",), 0),       # ID doppelt
            ("modObjekte", "ObjektDuplizieren", ("OBJ-999", "OBJ-006"), 0),  # Quelle fehlt
            ("modObjekte", "ObjektEntfernen", ("OBJ-002",), FEHLT),
            ("modObjekte", "ObjektAnlegen", ("OBJ-005",), 2),
            ("modObjekte", "ObjektPosition", ("OBJ-003",), 3),
            ("modObjekte", "AnnahmenWiederherstellen", (), 0),  # Entfernen stellt selbst her
            ("modObjekte", "LeereBloeckeAusblenden", (True,), FEHLT),
            ("modObjekte", "LeereBloeckeAusgeblendet", (), True),
            ("modObjekte", "LeereBloeckeAusblenden", (False,), FEHLT),
            ("modObjekte", "LeereBloeckeAusgeblendet", (), False),
            ("modObjekte", "LeereBloeckeAusblenden", (True,), FEHLT),
            ("modPruefung", "AnzahlFehler", (), 1),
            ("modPruefung", "AnzahlWarnungen", (), 0),
        ]),
        ("Etappe 9: Makros Prüfung und Variante festhalten", Modell(objekte=[obj]), [
            ("var_Bezeichnung", 0, "dritte"),
            ("var_Bezeichnung", 1, None),        # nach dem Leeren nur eine Zeile
            ("var_A", 0, Wie("vg_A", 0)),        # Endvermögen nach latenter Steuer
            ("var_B", 0, Wie("vg_B", 0)),
            ("var_C", 0, Wie("vg_C", 0)),
            ("var_Baseline", 0, Wie("vg_Baseline", 0)),
            ("var_DiffC", 0, Wie("vg_DiffC", 0)),
            ("var_Steuer", 0, Wie("vg_A", 14)),
            ("var_Pruefung", 0, "OK"),
        ], [
            ("modPruefung", "AnzahlFehler", (), 0),
            ("modPruefung", "AnzahlWarnungen", (), 0),
            ("modPruefung", "Auffaelligkeiten", (), ""),
            ("modVarianten", "VarianteFesthalten", ("erste",), 1),
            ("modVarianten", "VarianteFesthalten", ("zweite",), 2),
            ("modVarianten", "VarianteFesthalten", ("",), 0),  # Bezeichnung fehlt
            ("modVarianten", "FreieVariante", (), 3),
            ("modVarianten", "VariantenLeeren", (), FEHLT),
            ("modVarianten", "VarianteFesthalten", ("dritte",), 1),
            ("modRechnen", "NeuBerechnen", (), FEHLT),
        ]),
        # ohne Miete bleiben Erhaltung und AfA: Liquidität in allen 20 Jahren negativ
        # Anlagenverzeichnis (Abschnitt 21): KSt 1 trifft die Stammdaten des Testobjekts,
        # KSt 2: Gebäude 24.000 p. a., Außenanlage 6.000 p. a. mit Buchwert 15.000 Ende 2025
        # (Rest 3.000 in 2028), im Bau 50.000 ohne AfA; Restbuchwert Ende 2026 =
        # 852.000 + 9.000 + 50.000 = 911.000. Zeilen im Blatt Anlagen in der Reihenfolge der
        # Vorlage ohne den Abgang: 6 G+B ohne KOST1, 7 KOST1 9 ohne Objekt, 8 degressiv, 9 Wertpapiere
        ("Anlagen: Stammdaten und AfA je Anlage aus dem Anlagenverzeichnis",
         anlagen_eingelesen(), [
            ("obj_ID", 0, "KSt 1"),
            ("obj_AKGebaeude", 0, 800_000),
            ("obj_AKGuB", 0, 200_000),
            ("obj_Kaufjahr", 0, 2007),
            ("obj_AfASatz", 0, 0.02),
            ("obj_Restbuchwert", 0, 480_000),
            ("obj_AfAJahr", 0, 16_000),
            ("obj_Status", 0, "OK"),
            # nur noch Verkehrswert, Anteil, Großmaßnahme Jahr und Betrag
            ("obj_Annahmen", 0, 4),
            # Baujahr = frühester Gebäudezugang, orange und Warnung bis zur Bestätigung
            ("obj_Baujahr", 0, 2007),
            ("obj_AnlBau", 0, 2007),
            ("obj_AnlBauOffen", 0, 1),
            ("obj_Baujahr", 1, 2012),           # Außenanlage 2018 ist später
            ("obj_AnlAbn", 0, 1),
            ("obj_AnlAfA", 0, 16_000),
            ("obj_AnlDiff", 0, 0),
            ("prg_AfA", prg(1, 2027), 16_000),
            ("prg_Buchwert", prg(1, 2046), 160_000),
            ("obj_ID", 1, "KSt 2"),
            ("obj_AKGebaeude", 1, 1_310_000),
            ("obj_AKGuB", 1, 400_000),
            ("obj_Kaufjahr", 1, 2012),        # Außenanlage 2018 und Anbau 2025 sind später
            ("obj_AfASatz", 1, 30_000 / 1_310_000),
            ("obj_Restbuchwert", 1, 911_000),
            ("obj_AfAJahr", 1, 30_000),
            ("obj_AnlAbn", 1, 3),
            ("obj_AnlAK", 1, 3),
            ("obj_AnlGuB", 1, 1),
            ("obj_AnlKauf", 1, 3),
            ("obj_AnlAfA", 1, 30_000),
            ("prg_AfA", prg(2, 2027), 30_000),
            ("prg_Buchwert", prg(2, 2027), 881_000),
            ("prg_AfA", prg(2, 2028), 27_000),   # Außenanlage nur noch 3.000
            ("prg_AfA", prg(2, 2029), 24_000),
            ("prg_Buchwert", prg(2, 2046), 911_000 - 30_000 - 27_000 - 18 * 24_000),
            ("prg_AfAHalten", prg(2, 2028), 27_000),
            ("anl_Status", 0, "OK"),
            ("anl_AfA", 3, 24_000),
            ("anl_BWBasis", 3, 852_000),
            ("anl_AfABasis", 4, 6_000),
            ("anl_AfAJahre", 4, 6_000),           # 2027
            ("anl_Gruppe", 5, "abnutzbar"),
            ("anl_AfA", 5, 0),                    # im Bau
            ("anl_Zugang", 5, None),
            ("anl_Status", 6, "ohne ObjektID"),
            ("anl_Status", 7, "ObjektID fehlt im Blatt Objekte"),
            ("anl_AfA", 8, 200),                  # degressiv 25 % von 800
            ("anl_BWBasis", 8, 600),
            ("anl_AfAJahre", 8, 150),
            ("anl_Status", 9, "nicht im Modell (Art)"),
            ("anl_Status", 10, None),
        ] + befund(annahmen=2, anlagen=3, baujahr=2, zins_grob=2)),
        ("Anlagen: Verkauf mit Buchwert aus dem Anlagenverzeichnis",
         Modell(objekte=anlagen_eingelesen().objekte, anlagen=anlagen_eingelesen().anlagen,
                verkaeufe=[Verkauf("KSt 2", 2028, preis=2_000_000, nutzung_6b="nein")]), [
            ("vk_BuchwertGeb", 0, 854_000),
            ("vk_AKGuB", 0, 400_000),
            ("vk_Gewinn", 0, 2_000_000 - 854_000 - 400_000),
            ("vk_Vorbesitz", 0, 16),
            ("prg_AfA", prg(2, 2028), 27_000),
            ("prg_AfA", prg(2, 2029), 0),            # nach dem Verkauf
            ("prg_AfAHalten", prg(2, 2029), 24_000),  # Baseline hält weiter
            ("prg_BuchwertHalten", prg(2, 2029), 830_000),
            ("obj_Kritisch", 1, 1),                    # nur noch der Verkehrswertanteil
        ] + befund(annahmen=2, anlagen=3, kritisch=1, baujahr=2, zins_grob=2,
                   zins_verkauf=1)),
        ("Anlagen: nur Aufschlüsselung der Abschreibungen aus dem Kostenstellenblatt",
         anlagen_eingelesen(inventar=False), [
            ("anl_Status", 0, "OK"),
            ("anl_Status", 4, None),                  # vier Gruppen
            ("obj_Restbuchwert", 1, 911_000),
            ("obj_AnlAK", 1, 0),
            ("obj_AKGebaeude", 1, 1_500_000),         # Annahme AfA lt. BWA / 2 %
            ("obj_Kaufjahr", 1, 2011),                # Annahme
            ("obj_AnlBau", 1, None),                  # Gruppen ohne AHK-Datum
            ("obj_AnlBauOffen", 1, 0),
            ("obj_Annahmen", 1, 9),                   # Restbuchwert kommt aus den Gruppen
            ("prg_AfA", prg(2, 2027), 30_000),
            ("prg_AfA", prg(2, 2028), 27_000),
            ("prg_AfA", prg(2, 2029), 24_000),
            ("prg_AfA", prg(1, 2027), 16_000),
        ]),
        # G+B ohne KOST1, über die Bezeichnung („Beispielweg 7“) KSt 2 zugeordnet
        ("Anlagen: Zuordnung über die Bezeichnung ohne KOST1",
         anlagen_eingelesen(ohne_kost1=[Anlage(nr="100020", bw_stand=10_000, art="G+B",
                                               methode="keine",
                                               bezeichnung="Stellplätze Beispielweg")]), [
            ("anl_ID", 10, "KSt 2"),
            ("anl_Zuordnung", 10, "Bezeichnung: beispielweg"),
            ("anl_Zuordnung", 0, "KOST1"),
            ("anl_Status", 10, "OK"),
            ("obj_AKGuB", 1, 410_000),
        ] + befund(annahmen=2, anlagen=3, anlagen_bez=1, baujahr=2, zins_grob=2)),
        # G+B als „Grund u. Boden Kostenstelle 2“ ohne KOST1: zählt zu AK G+B (400.000 + 10.000)
        # und mindert beim Verkauf den Gewinn auf G+B: 2 Mio × 30 % − 410.000 = 190.000
        ("Anlagen: G+B über „Kostenstelle“ in der Bezeichnung, wirkt im Verkauf",
         dataclasses.replace(
             anlagen_eingelesen(ohne_kost1=[Anlage(nr="100021", bw_stand=10_000, art="G+B",
                                                   methode="keine",
                                                   bezeichnung="Grund u. Boden Kostenstelle 2")]),
             verkaeufe=[Verkauf("KSt 2", 2027, preis=2_000_000, anteil_gub=0.3)]), [
            ("anl_ID", 10, "KSt 2"),
            ("anl_Zuordnung", 10, "Bezeichnung: KSt 2"),
            ("obj_AKGuB", 1, 410_000),
            ("vk_AKGuB", 0, 410_000),
            ("vk_ErloesGuB", 0, 600_000),
            ("vk_GewinnGuB", 0, 190_000),
        ]),
        ("Anlagen: Baujahr eingetippt bestätigt, keine Warnung mehr", baujahr_bestaetigt(), [
            ("obj_Baujahr", 0, 1990),
            ("obj_AnlBauOffen", 0, 0),
            ("obj_AnlBauOffen", 1, 1),
        ] + befund(annahmen=2, anlagen=3, baujahr=1, zins_grob=2)),
        ("Anlagen: Stand des Anlagenverzeichnisses gleich Basisjahr",
         anlagen_eingelesen(par_AnlStand=2026), [
            ("obj_Restbuchwert", 0, 496_000),
            ("obj_AnlAfA", 0, 16_000),                # AfA im Stand-Jahr lt. Inventar
            ("prg_AfA", prg(1, 2027), 16_000),
            ("prg_Buchwert", prg(1, 2027), 480_000),
            ("obj_Restbuchwert", 1, 941_000),
            ("prg_AfA", prg(2, 2028), 30_000),
            ("prg_AfA", prg(2, 2029), 27_000),
        ]),
        ("Etappe 9: Makro Prüfung meldet Fehler und Hinweis", Modell(objekte=[ohne_miete]), [
            ("pr_Gesamt", 0, "1 Fehler"),
        ], [
            ("modPruefung", "AnzahlFehler", (), 1),
            ("modPruefung", "Auffaelligkeiten", (),
             f"{FEHLER}: {PRUEFUNGEN[0].bezeichnung} (1)\n"
             f"{HINWEIS}: {PRUEFUNGEN[pr('liquiditaet')].bezeichnung} (20)\n"),
        ]),
    ]


def eingeklappt(ws, spalte: int) -> tuple:
    """Gliederungsebene und Ausblendung einer Spalte (Spaltenbereiche mit min/max beachten)."""
    for dim in ws.column_dimensions.values():
        if dim.min and dim.max and dim.min <= spalte <= dim.max:
            return dim.outline_level, bool(dim.hidden)
    return 0, False


def auswahl(ws, zelle: str) -> tuple:
    """Auswahlliste der Zelle (formula1 einer Listenüberprüfung) und Anzahl ihrer Überprüfungen."""
    treffer = [dv for dv in ws.data_validations.dataValidation if zelle in dv.sqref]
    listen = [dv.formula1 for dv in treffer if dv.type == "list"]
    return (listen[0] if listen else None), len(treffer)


def ueberlappungen(wb) -> list:
    """Zellbereiche mit mehr als einer Datenüberprüfung; Excel verwirft solche Überprüfungen."""
    fundstellen = []
    for ws in wb:
        bereiche = [(i, r) for i, dv in enumerate(ws.data_validations.dataValidation)
                    for r in dv.sqref.ranges]
        for a, (i, r) in enumerate(bereiche):
            for j, q in bereiche[a + 1:]:
                if i != j and not r.isdisjoint(q):
                    fundstellen.append(f"{ws.title}!{r.coord}")
    return fundstellen


def pruefe(fall: str, wb, pruefungen) -> int:
    """Sollwerte eines Falls prüfen, Ergebnis ausgeben, Anzahl Abweichungen zurückgeben."""
    fehler = 0
    for name, zeile, soll in pruefungen:
        if name == AUSGEBLENDET:             # Zeile im Blatt Prognose ausgeblendet?
            ist = bool(wb["Prognose"].row_dimensions[zeile].hidden)
        elif name.startswith(EINGEKLAPPT):
            ist = eingeklappt(wb[name[len(EINGEKLAPPT):]], zeile)
        elif name.startswith(ZEILE_EINGEKLAPPT):
            dim = wb[name[len(ZEILE_EINGEKLAPPT):]].row_dimensions[zeile]
            ist = (dim.outline_level, bool(dim.hidden))
        elif name.startswith(AUSWAHL):
            ist = auswahl(getattr(wb, "erzeugt", wb)[name[len(AUSWAHL):]], zeile)
        elif name == UEBERLAPPT:
            ist = ueberlappungen(getattr(wb, "erzeugt", wb))
        elif name.startswith(KOPF):
            blatt, kopfzeile = name[len(KOPF):].rsplit(":", 1)
            c = wb[blatt].cell(row=int(kopfzeile), column=zeile)
            ist = f"{c.value}{' [Tooltip]' if c.comment else ''}"
        else:
            ist = wert(wb, name, zeile)
        if ist == "":
            ist = None
        if isinstance(soll, Wie):
            soll = wert(wb, soll.name, soll.zeile)
        if gleich(ist, soll):
            print(f"OK      {fall}: {name}[{zeile}] = {ist!r}")
        else:
            fehler += 1
            print(f"FEHLER  {fall}: {name}[{zeile}] Soll {soll!r}, Ist {ist!r}")
    return fehler


def makro_lauf(lo, fall: str, modell: Modell, aufrufe, ordner: Path, zellen: dict = None):
    """Etappe 9: Mappe als .xlsm bauen, Makros in LibreOffice ausführen, Ergebnis lesen.

    Liefert die neu berechnete Mappe (Werte) und die Anzahl abweichender Rückgabewerte.
    """
    ordner.mkdir(parents=True, exist_ok=True)
    xlsm = speichere_mit_makros(mappe(modell, zellen), ordner / "mappe.xlsm", lo)
    fehler = 0
    doc = lo.laden(xlsm, makros=True)
    try:
        for modul, prozedur, argumente, soll in aufrufe:
            aufruf = f"{modul}.{prozedur}{argumente!r}"
            try:
                ist = lo.makro(doc, modul, prozedur, *argumente)
            except Exception as e:  # Makro nicht gefunden oder Laufzeitfehler in UNO
                fehler += 1
                print(f"FEHLER  {fall}: {aufruf} bricht ab: {str(e).splitlines()[0]}")
                continue
            if soll is FEHLT:
                print(f"OK      {fall}: {aufruf}")
            elif gleich(ist, soll):
                print(f"OK      {fall}: {aufruf} -> {ist!r}")
            else:
                fehler += 1
                print(f"FEHLER  {fall}: {aufruf} Soll {soll!r}, Ist {ist!r}")
        lo.speichern(doc, ordner / "gerechnet.xlsx")
    finally:
        doc.close(True)
    return load_workbook(ordner / "gerechnet.xlsx", data_only=True), fehler


def lauf(i: int, tmp: Path, lo_stapel: contextlib.ExitStack = None) -> tuple:
    """Fall i rechnen und prüfen. Liefert (Fallname, Ausgabe, Anzahl Abweichungen)."""
    fall, objekte, pruefungen, *aufrufe = _faelle()[i]
    zellen = None
    if isinstance(objekte, Getippt):
        objekte, zellen = objekte.modell, objekte.zellen
    modell = objekte if isinstance(objekte, Modell) else Modell(objekte=objekte)
    if not fall.startswith("Erhaltung"):
        modell.parameter = {**OHNE_ALTERUNG, **modell.parameter}
    if fall not in VOLLE_GROESSE:
        modell.kapazitaet = TEST_KAPAZITAET
    fehler = 0
    ausgabe = io.StringIO()
    with contextlib.redirect_stdout(ausgabe):
        if aufrufe:
            lo = _libreoffice(tmp, lo_stapel)
            wb, fehler = makro_lauf(lo, fall, modell, aufrufe[0], tmp / f"fall{i}", zellen)
        else:
            wb = durchrechnen(modell, tmp / f"fall{i}", zellen, tmp / "lo-profil")
        fehler += pruefe(fall, wb, pruefungen)
    return fall, ausgabe.getvalue(), fehler


_FAELLE = None
_LO = None


def _faelle() -> list:
    global _FAELLE
    if _FAELLE is None:
        _FAELLE = list(faelle())
    return _FAELLE


def _libreoffice(tmp: Path, stapel: contextlib.ExitStack):
    """Eine LibreOffice-Sitzung (UNO) für alle Makrofälle eines Prozesses."""
    global _LO
    if _LO is None:
        _LO = stapel.enter_context(LibreOffice(tmp / "makros"))
    return _LO


def main() -> int:
    """Alle Fälle prüfen; optional nur Fälle, deren Name den Filtertext enthält."""
    argumente = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    argumente.add_argument("filter", nargs="?", default="", help="Teil des Fallnamens")
    argumente.add_argument("-j", "--prozesse", type=int, default=os.cpu_count() or 1,
                           help="Fälle parallel (Standard: Anzahl Kerne)")
    args = argumente.parse_args()
    auswahl = [i for i, f in enumerate(_faelle()) if args.filter in f[0]]
    # Makrofälle teilen sich eine UNO-Sitzung im Hauptprozess, die übrigen laufen im Pool
    makro = [i for i in auswahl if len(_faelle()[i]) > 3]
    rest = [i for i in auswahl if i not in makro]
    fehler = 0
    fertig = 0

    def melden(fall: str, ausgabe: str, n: int) -> None:
        nonlocal fehler, fertig
        fehler += n
        fertig += 1
        print(ausgabe, end="")
        print(f"[{fertig}/{len(auswahl)}] {'FEHLER' if n else 'OK'}  {fall}", flush=True)

    with tempfile.TemporaryDirectory() as tmp, contextlib.ExitStack() as stapel:
        tmp = Path(tmp)
        prozesse = max(1, min(args.prozesse, len(rest)))
        pool = stapel.enter_context(concurrent.futures.ProcessPoolExecutor(
            prozesse, initializer=_arbeiter_start, initargs=(str(tmp),)))
        laeufe = [pool.submit(_im_arbeiter, i) for i in rest]
        for i in makro:
            melden(*lauf(i, tmp, stapel))
        for erledigt in concurrent.futures.as_completed(laeufe):
            melden(*erledigt.result())
    print(f"\n{fehler} Abweichung(en)" if fehler else "\nAlle Prüfungen bestanden.")
    return 1 if fehler else 0


_ORDNER = None


def _arbeiter_start(tmp: str) -> None:
    """Je Arbeitsprozess ein eigener Ordner mit eigenem LibreOffice-Profil."""
    global _ORDNER
    _ORDNER = Path(tempfile.mkdtemp(prefix=f"prozess{os.getpid()}_", dir=tmp))
    _faelle()


def _im_arbeiter(i: int) -> tuple:
    return lauf(i, _ORDNER)


if __name__ == "__main__":
    sys.exit(main())
