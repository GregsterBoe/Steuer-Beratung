"""Prüfskript: Mappe generieren, mit LibreOffice headless durchrechnen, gegen Sollwerte prüfen.

Aufruf: python -m pruefung.pruefen [Teil des Fallnamens]
Beendet sich mit Fehlercode 1, sobald ein Fall abweicht.

Fälle mit Makroaufrufen (Etappe 9) werden als .xlsm gebaut, in LibreOffice mit
Makros geöffnet, die Makros ausgeführt und das Ergebnis danach geprüft.
"""

import contextlib
import dataclasses
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from prognosemodell.einlesen import lese_kostenstellen, zusammenfuehren
from prognosemodell.makros import LibreOffice, speichere_mit_makros
from prognosemodell.mappe import erstelle_mappe
from prognosemodell.modelle import (FEHLER, HINWEIS, MAX_OBJEKTE, PRUEFUNGEN,
                                    STATUS_ANNAHME_GELOESCHT, WARNUNG, Modell, Neuobjekt,
                                    Objekt, Verkauf, prognosejahre)
from prognosemodell.testdaten import testobjekt
from prognosemodell.vorlagen import erstelle_vorlage

TOLERANZ = 0.01  # ein Cent
# Fälle vor der Alterslogik: Sollwerte ohne Alterung, Anlaufminderung und Großmaßnahmen.
# Gilt für jeden Fall, dessen Name nicht mit „Erhaltung“ beginnt.
OHNE_ALTERUNG = {"par_ErhAlterung": 0, "par_NeuErhAnlaufFaktor": 1, "par_SanQuote": 0}
FEHLT = object()        # Makro ohne Rückgabewert: nur prüfen, dass es fehlerfrei läuft
AUSGEBLENDET = "Zeile ausgeblendet"  # statt Bereichsname: Zeilennummer im Blatt Prognose


@dataclasses.dataclass(frozen=True)
class Wie:
    """Sollwert = Wert eines anderen Bereichs derselben Mappe."""
    name: str
    zeile: int = 0


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
    return prg(MAX_OBJEKTE + neu_nr, jahr)


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
def vorlage_eingelesen() -> Modell:
    """Modell aus der BWA-Vorlage: KSt 1 mit den Stammdaten des Testobjekts, KSt 2 ohne."""
    with tempfile.TemporaryDirectory() as tmp:
        pfad = Path(tmp) / "vorlage.xlsx"
        erstelle_vorlage().save(pfad)
        laufende, _ = lese_kostenstellen(pfad, 2026)
    stamm = [dataclasses.replace(testobjekt(), objekt_id="KSt 1", name=None)]
    return Modell(objekte=zusammenfuehren(stamm, laufende),
                  kostenstellen={lw.objekt_id: lw for lw in laufende})


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
            ("SB1:Kapitalanlage (Nettoerlös − Reinvestition − Steuer ca.)", 1, 200_000),
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
            ("KAUF:= Cash Flow (vor Finanzierung)", 1, 48_900),
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
         Modell(objekte=[dataclasses.replace(obj, baujahr=1980)],
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
            (AUSGEBLENDET, 2 + MAX_OBJEKTE * 20, True),  # leerer Neuobjektblock
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
        ("Etappe 9: Makro Prüfung meldet Fehler und Hinweis", Modell(objekte=[ohne_miete]), [
            ("pr_Gesamt", 0, "1 Fehler"),
        ], [
            ("modPruefung", "AnzahlFehler", (), 1),
            ("modPruefung", "Auffaelligkeiten", (),
             f"{FEHLER}: {PRUEFUNGEN[0].bezeichnung} (1)\n"
             f"{HINWEIS}: {PRUEFUNGEN[pr('liquiditaet')].bezeichnung} (20)\n"),
        ]),
    ]


def pruefe(fall: str, wb, pruefungen) -> int:
    """Sollwerte eines Falls prüfen, Ergebnis ausgeben, Anzahl Abweichungen zurückgeben."""
    fehler = 0
    for name, zeile, soll in pruefungen:
        if name == AUSGEBLENDET:             # Zeile im Blatt Prognose ausgeblendet?
            ist = bool(wb["Prognose"].row_dimensions[zeile].hidden)
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


def makro_lauf(lo, fall: str, modell: Modell, aufrufe, ordner: Path):
    """Etappe 9: Mappe als .xlsm bauen, Makros in LibreOffice ausführen, Ergebnis lesen.

    Liefert die neu berechnete Mappe (Werte) und die Anzahl abweichender Rückgabewerte.
    """
    ordner.mkdir(parents=True, exist_ok=True)
    xlsm = speichere_mit_makros(erstelle_mappe(modell), ordner / "mappe.xlsm", lo)
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


def main() -> int:
    """Alle Fälle prüfen; optional nur Fälle, deren Name den ersten Aufrufparameter enthält."""
    filter_ = sys.argv[1] if len(sys.argv) > 1 else ""
    fehler = 0
    with tempfile.TemporaryDirectory() as tmp, contextlib.ExitStack() as stapel:
        lo = None
        for i, (fall, objekte, pruefungen, *aufrufe) in enumerate(faelle()):
            if filter_ not in fall:
                continue
            modell = objekte if isinstance(objekte, Modell) else Modell(objekte=objekte)
            if not fall.startswith("Erhaltung"):
                modell.parameter = {**OHNE_ALTERUNG, **modell.parameter}
            if aufrufe:
                if lo is None:                   # eine LibreOffice-Sitzung für alle Makrofälle
                    lo = stapel.enter_context(LibreOffice(Path(tmp) / "makros"))
                wb, n = makro_lauf(lo, fall, modell, aufrufe[0], Path(tmp) / f"fall{i}")
                fehler += n
            else:
                wb = durchrechnen(modell, Path(tmp) / f"fall{i}")
            fehler += pruefe(fall, wb, pruefungen)
    print(f"\n{fehler} Abweichung(en)" if fehler else "\nAlle Prüfungen bestanden.")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main())
