"""Prüfskript: Mappe generieren, mit LibreOffice headless durchrechnen, gegen Sollwerte prüfen.

Aufruf: python -m pruefung.pruefen
Beendet sich mit Fehlercode 1, sobald ein Fall abweicht.
"""

import dataclasses
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from prognosemodell.mappe import erstelle_mappe
from prognosemodell.modelle import MAX_OBJEKTE, Modell, Neuobjekt, Verkauf, prognosejahre
from prognosemodell.testdaten import testobjekt

TOLERANZ = 0.01  # ein Cent


def durchrechnen(modell: Modell, arbeitsordner: Path):
    """Mappe schreiben, per LibreOffice neu berechnen lassen, Werte zurückgeben."""
    roh = arbeitsordner / "roh" / "mappe.xlsx"
    roh.parent.mkdir(parents=True, exist_ok=True)
    erstelle_mappe(modell).save(roh)
    aus = arbeitsordner / "gerechnet"
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("LibreOffice (soffice) nicht gefunden.")
    subprocess.run(
        [soffice, f"-env:UserInstallation=file://{arbeitsordner}/lo-profil", "--headless",
         "--calc", "--convert-to", "xlsx", "--outdir", str(aus), str(roh)],
        check=True, capture_output=True, timeout=120,
    )
    return load_workbook(aus / "mappe.xlsx", data_only=True)


def wert(wb, name: str, zeile_im_bereich: int = 0):
    """Wert eines benannten Bereichs (bei Spaltenbereichen: n-te Zeile)."""
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    ref = ref.replace("$", "").split(":")[0]
    zelle = wb[blatt][ref]
    return wb[blatt].cell(row=zelle.row + zeile_im_bereich, column=zelle.column).value


def gleich(ist, soll) -> bool:
    if isinstance(soll, (int, float)) and isinstance(ist, (int, float)):
        return abs(ist - soll) <= TOLERANZ
    return ist == soll


def prg(objekt_nr: int, jahr: int) -> int:
    """Zeile im Prognosebereich für Objekt objekt_nr (1 = erstes) und Jahr."""
    return (objekt_nr - 1) * prognosejahre() + jahr - 2027


def prg_neu(neu_nr: int, jahr: int) -> int:
    """Zeile im Prognosebereich für Neuobjekt neu_nr; die Neuobjekte folgen auf alle Objektblöcke."""
    return prg(MAX_OBJEKTE + neu_nr, jahr)


def rls(jahr: int) -> int:
    """Zeile im Rücklagenspiegel je Jahr; erste Zeile ist das erste Prognosejahr."""
    return jahr - 2027


def lj(jahr: int) -> int:
    """Zeile in den Jahrestabellen von Liquidität und Auswertung; erste Zeile ist 2027."""
    return jahr - 2027


def ueb(jahr: int) -> int:
    """Zeile im Übersichtsbereich; erste Zeile ist das Basisjahr."""
    return jahr - 2026


# Je Fall: Name, Objekte (oder ein Modell mit Verkäufen), Liste von (benannter Bereich, Zeile im Bereich, Sollwert)
# Sollwerte der Etappe 2 von Hand: AfA 800.000 × 2 % = 16.000, Miete 60.000 × 1,02^n,
# Erhaltung 8.000 × 1,025^n, weitere Ausgaben × 1,02^n, n = Jahr − 2026.
# Übersicht: Verkehrswert 1.400.000 × 1,02^n je Objekt.
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
    abnahme4 = dataclasses.replace(obj, restbuchwert=496_000)  # Buchwert Ende 2027: 480.000
    ohne_quote = dataclasses.replace(obj, objekt_id="OBJ-005", vk_quote_gebaeude=None)
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
            ("ueb_OhneWert", 0, 0),
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
        ("Verkäufe: Statusprüfung, fehlender Verkehrswert",
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
            ("ueb_OhneWert", 0, 1),
            ("ueb_Baseline", ueb(2026), 2_800_000),  # OBJ-003 zählt mit 0
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
                verkaeufe=[Verkauf("OBJ-001", 2028),
                           Verkauf("OBJ-005", 2028, preis=1_000_000),
                           Verkauf("OBJ-004", 2029, preis=1_000_000, nutzung_6b="ja")]), [
            ("vk_Status", 0, "Verkaufspreis fehlt"),
            ("vk_Status", 1, "Aufteilung fehlt: Anteil G+B oder Verkehrswertanteil"),
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
         [obj, obj2], [
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
                                                  nutzung_6b="nein")]), [
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
                verkaeufe=[Verkauf("OBJ-001", 2027, preis=1_400_000, nutzung_6b="ja")]), [
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
                                      quelle="RL-OBJ-001")]), [
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
                                                  anteil_gub=0.8, nutzung_6b="ja")]), [
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
    ]


def main() -> int:
    fehler = 0
    with tempfile.TemporaryDirectory() as tmp:
        for i, (fall, objekte, pruefungen) in enumerate(faelle()):
            modell = objekte if isinstance(objekte, Modell) else Modell(objekte=objekte)
            wb = durchrechnen(modell, Path(tmp) / f"fall{i}")
            for name, zeile, soll in pruefungen:
                ist = wert(wb, name, zeile)
                if ist == "":
                    ist = None
                if gleich(ist, soll):
                    print(f"OK      {fall}: {name}[{zeile}] = {ist!r}")
                else:
                    fehler += 1
                    print(f"FEHLER  {fall}: {name}[{zeile}] Soll {soll!r}, Ist {ist!r}")
    print(f"\n{fehler} Abweichung(en)" if fehler else "\nAlle Prüfungen bestanden.")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main())
