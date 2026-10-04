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
from prognosemodell.formeln import NEU_SZENARIO_B, RUECKLAGE_SZENARIO_B
from prognosemodell.modelle import (MAX_OBJEKTE, STEUERWELTEN, VERGLEICH_KENNZAHLEN, Modell,
                                    Neuobjekt, Verkauf)
from prognosemodell.testdaten import testobjekt

TOLERANZ = 0.01  # ein Cent
STARTJAHR, JAHRE = 2027, 20


def pz(objekt: int, jahr: int) -> int:
    """Zeile im Prognosebereich für die n-te Objektzeile (ab 0) und ein Jahr."""
    return objekt * JAHRE + (jahr - STARTJAHR)


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


def vg(key: str) -> int:
    """Zeile einer Kennzahl in den Bereichen vg_* des Blatts Vergleich."""
    return next(i for i, k in enumerate(VERGLEICH_KENNZAHLEN) if k.key == key)


def gleich(ist, soll) -> bool:
    if isinstance(soll, (int, float)) and isinstance(ist, (int, float)):
        return abs(ist - soll) <= TOLERANZ
    return ist == soll


# Je Fall: Name, Modell, Liste von (benannter Bereich, Zeile im Bereich, Sollwert)
def faelle():
    obj = testobjekt()
    ohne_miete = dataclasses.replace(obj, miete=None)
    zu_hoch = dataclasses.replace(obj, restbuchwert=900_000)
    spaet = dataclasses.replace(obj, kaufjahr=2030)
    zweites = dataclasses.replace(obj, name="Duplikat")
    # 50.000 Restbuchwert: 2027 und 2028 volle AfA, 2029 Rest 10.000, ab 2030 null
    kurz = dataclasses.replace(obj, objekt_id="OBJ-002", restbuchwert=50_000)
    miete_2046 = 60_000 * 1.02 ** 20
    afa = 20_000
    erh_2030 = 8_000 * 1.025 ** 4
    # Abnahmefall Projektplan: Restbuchwert 500.000, Ende 2027 Gebäude 480.000 + G+B 200.000
    vk_obj = dataclasses.replace(obj, restbuchwert=500_000)
    vk_2027 = Verkauf("OBJ-001", 2027, preis=1_400_000, kosten=0, aufteilung="Buchwert",
                      nutzung_6b="ja")
    erloes_geb_bw = 1_400_000 * 480_000 / 680_000
    # Verkauf 2030 über Mietfaktor 20, Kosten 30.000, Aufteilung aus dem Parameterblatt
    miete_2030 = 60_000 * 1.02 ** 4
    vk_2030 = Verkauf("OBJ-001", 2030, faktor=20, kosten=30_000, nutzung_6b="nein")
    netto_2030 = 20 * miete_2030 - 30_000
    ohne_quote = dataclasses.replace(obj, vk_quote_gebaeude=None)

    # Etappe 5: Rücklagenspiegel, Jahreszeile im Spiegel ab 2027 = Index 0
    gewinn_geb_bw = erloes_geb_bw - 480_000
    gewinn_gub_bw = 1_400_000 - erloes_geb_bw - 200_000
    zuschlag = 720_000 * 0.06 * 4
    gewinn_2030 = netto_2030 - 520_000

    def rj(jahr: int) -> int:
        return jahr - STARTJAHR

    def vk_status(verkauf, objekte=(obj,)):
        return Modell(list(objekte), verkaeufe=[verkauf])

    # Etappe 6: Reinvestition; Neuobjekte stehen in der Prognose hinter allen Bestandsblöcken
    def nz(n: int, jahr: int) -> int:
        return pz(MAX_OBJEKTE + n, jahr)

    def neu(neu_id="NEU-001", kaufjahr=2028, kaufpreis=900_000, anteil_gub=0.0, nebenkosten=0,
            quelle="R-OBJ-001"):
        return Neuobjekt(neu_id, "Neubau", kaufjahr, kaufpreis, anteil_gub, nebenkosten,
                         afa_satz=0.03, mietrendite=0.05, erhaltungsquote=0.01, quelle=quelle)

    def reinvest(*neuobjekte, verkauf=vk_2027, objekte=(vk_obj,)):
        return Modell(list(objekte), verkaeufe=[verkauf], neuobjekte=list(neuobjekte))

    rest_teil = 720_000 - 400_000
    basis_zweit = 800_000 - (gewinn_geb_bw - 300_000)

    # Etappe 7: Jahresblätter, Zeile 0 = 2027, Zeile 20 = Summe über alle Jahre
    summe = JAHRE
    ergebnisse = [60_000 * 1.02 ** k - 8_000 * 1.025 ** k - afa for k in range(1, JAHRE + 1)]
    ergebnis_2028 = ergebnisse[1]
    ergebnis_2030 = miete_2030 - erh_2030 - afa
    steuer_2030 = (ergebnis_2030 + gewinn_2030) * 0.30
    fluss_2027 = 53_000 - 9_900
    fluss_2028 = 60_000 * 1.02 ** 2 - 8_000 * 1.025 ** 2 - ergebnis_2028 * 0.30

    # Etappe 8: Szenariovergleich am Abnahmefall Reinvestition, Alternativrendite 0
    szenario_a = reinvest(neu())
    szenario_a.parameter = {"par_Alternativrendite": 0}
    szenario_b = dataclasses.replace(szenario_a, parameter={"par_Alternativrendite": 0,
                                                            "par_Szenario": "B"})
    neu_fluesse = []
    for k in range(18):                  # Neuobjekt 2029 bis 2046
        m, e = 45_000 * 1.02 ** k, 9_000 * 1.025 ** k
        neu_fluesse.append(m - e - (m - e - 5_400) * 0.30)
    vw_neu = 900_000 * 1.02 ** 18
    bw_neu = 180_000 - 18 * 5_400
    anlage_a = 53_000 + 1_400_000 - 9_900 - 900_000 + sum(neu_fluesse)
    latent_a = (vw_neu - bw_neu) * 0.30
    vermoegen_a = vw_neu + anlage_a - latent_a
    steuer_b_2027 = (33_000 + 720_000) * 0.30
    anlage_b = 53_000 + 1_400_000 - steuer_b_2027
    assert vermoegen_a > anlage_b        # Abnahme: ohne Alternativrendite liegt A vorn
    # Szenario B mit 4 % Alternativrendite: 2,8 % nach Steuern auf den Bestand Ende 2027
    zins_b = sum(anlage_b * 1.028 ** n * 0.04 for n in range(19))

    return [
        ("Etappe 1: Stammdaten vollständig", Modell([obj]), [
            ("par_Startjahr", 0, 2027),
            ("par_Endjahr", 0, 2046),
            ("par_StatusSteuerwelt", 0, "OK"),
            ("obj_ID", 0, "OBJ-001"),
            ("obj_Status", 0, "OK"),
            ("obj_Status", 1, None),  # leere Zeile bleibt leer
        ]),
        ("Etappe 1: Pflichtfeld fehlt", Modell([ohne_miete]),
         [("obj_Status", 0, "Pflichtfeld fehlt")]),
        ("Etappe 1: Restbuchwert zu hoch", Modell([zu_hoch]),
         [("obj_Status", 0, "Restbuchwert > AK Gebäude")]),
        ("Etappe 1: Kaufjahr nach Basisjahr", Modell([spaet]),
         [("obj_Status", 0, "Kaufjahr nach Basisjahr")]),
        ("Etappe 1: ObjektID doppelt", Modell([obj, zweites]),
         [("obj_Status", 0, "ObjektID doppelt"), ("obj_Status", 1, "ObjektID doppelt")]),
        ("Etappe 2: AfA-Ende 2046", Modell([obj]), [
            ("prg_ID", pz(0, 2027), "OBJ-001"),
            ("prg_Jahr", pz(0, 2027), 2027),
            ("prg_Jahr", pz(0, 2046), 2046),
            ("prg_Aktiv", pz(0, 2027), 1),
            ("prg_AfA", pz(0, 2027), 20_000),
            ("prg_Buchwert", pz(0, 2027), 380_000),
            ("prg_AfA", pz(0, 2046), 20_000),
            ("prg_Buchwert", pz(0, 2045), 20_000),
            ("prg_Buchwert", pz(0, 2046), 0),
            ("prg_Miete", pz(0, 2027), 61_200),
            ("prg_Erhaltung", pz(0, 2027), 8_200),
            ("prg_Ergebnis", pz(0, 2027), 61_200 - 8_200 - 20_000),
            ("prg_Miete", pz(0, 2046), miete_2046),
            ("prg_ID", pz(1, 2027), None),       # leerer Block bleibt leer
            ("prg_Aktiv", pz(1, 2027), 0),
            ("prg_Ergebnis", pz(1, 2027), 0),
        ]),
        ("Etappe 2: AfA-Stopp im Raster (zweites Objekt)", Modell([obj, kurz]), [
            ("prg_ID", pz(1, 2027), "OBJ-002"),
            ("prg_AfA", pz(1, 2028), 20_000),
            ("prg_AfA", pz(1, 2029), 10_000),
            ("prg_Buchwert", pz(1, 2029), 0),
            ("prg_AfA", pz(1, 2030), 0),
            ("prg_Buchwert", pz(1, 2046), 0),
            ("prg_Erhaltung", pz(1, 2030), erh_2030),
            ("prg_Ergebnis", pz(1, 2030), 60_000 * 1.02 ** 4 - erh_2030),
            ("prg_AfA", pz(0, 2030), 20_000),    # erstes Objekt unberührt
        ]),
        ("Etappe 2: Objekt mit Fehlerstatus rechnet nicht", Modell([ohne_miete]), [
            ("prg_Aktiv", pz(0, 2027), 0),
            ("prg_AfA", pz(0, 2027), 0),
            ("prg_Ergebnis", pz(0, 2027), 0),
        ]),
        ("Etappe 3: Indexierung mit Standardraten", Modell([obj]), [
            ("prg_Miete", pz(0, 2028), 60_000 * 1.02 ** 2),
            ("prg_Miete", pz(0, 2036), 60_000 * 1.02 ** 10),
            ("prg_Erhaltung", pz(0, 2028), 8_000 * 1.025 ** 2),
            ("prg_Erhaltung", pz(0, 2046), 8_000 * 1.025 ** 20),
            ("prg_Ergebnis", pz(0, 2046), miete_2046 - 8_000 * 1.025 ** 20 - afa),
        ]),
        ("Etappe 3: geänderte Raten (Miete 3 %, Erhaltung 0 %)",
         Modell([obj], {"par_Mietsteig": 0.03, "par_Erhaltsteig": 0}), [
            ("prg_Miete", pz(0, 2027), 61_800),
            ("prg_Miete", pz(0, 2036), 60_000 * 1.03 ** 10),
            ("prg_Erhaltung", pz(0, 2027), 8_000),
            ("prg_Erhaltung", pz(0, 2046), 8_000),
            ("prg_Ergebnis", pz(0, 2036), 60_000 * 1.03 ** 10 - 8_000 - afa),
        ]),
        ("Etappe 3: Raten null und negativ",
         Modell([obj], {"par_Mietsteig": 0, "par_Erhaltsteig": -0.01}), [
            ("prg_Miete", pz(0, 2027), 60_000),
            ("prg_Miete", pz(0, 2046), 60_000),
            ("prg_Erhaltung", pz(0, 2027), 7_920),
            ("prg_Erhaltung", pz(0, 2046), 8_000 * 0.99 ** 20),
        ]),
        ("Etappe 3: Basisjahr verschoben (2030)", Modell([obj], {"par_Basisjahr": 2030}), [
            ("prg_Jahr", pz(0, 2027), 2031),
            ("prg_Miete", pz(0, 2027), 61_200),   # Index zählt ab Basisjahr, nicht ab 2026
            ("prg_Erhaltung", pz(0, 2027), 8_200),
        ]),
        ("Etappe 4: Verkauf mit Gewinn, Aufteilung nach Buchwert",
         Modell([vk_obj], verkaeufe=[vk_2027]), [
            ("vk_Status", 0, "OK"),
            ("vk_Status", 1, None),              # leere Zeile bleibt leer
            ("vk_PreisAngesetzt", 0, 1_400_000),
            ("vk_BuchwertGeb", 0, 480_000),
            ("vk_BuchwertGesamt", 0, 680_000),
            ("vk_Gewinn", 0, 720_000),
            ("vk_ErloesGeb", 0, erloes_geb_bw),
            ("vk_GewinnGeb", 0, erloes_geb_bw - 480_000),
            ("vk_GewinnGuB", 0, 1_400_000 - erloes_geb_bw - 200_000),
            ("prg_Aktiv", pz(0, 2027), 1),       # Verkaufsjahr rechnet noch voll
            ("prg_Ergebnis", pz(0, 2027), 61_200 - 8_200 - 20_000),
            ("prg_Aktiv", pz(0, 2028), 0),
            ("prg_Miete", pz(0, 2028), 0),
            ("prg_AfA", pz(0, 2028), 0),
            ("prg_Buchwert", pz(0, 2028), 0),
            ("prg_Ergebnis", pz(0, 2046), 0),
        ]),
        ("Etappe 4: hälftige Aufteilung nach Verkehrswert",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, aufteilung="Verkehrswert")]), [
            ("vk_ErloesGeb", 0, 700_000),
            ("vk_GewinnGeb", 0, 220_000),
            ("vk_GewinnGuB", 0, 500_000),
            ("vk_Gewinn", 0, 720_000),
        ]),
        ("Etappe 4: Verkauf 2030 über Mietfaktor, mit Kosten", Modell([obj], verkaeufe=[vk_2030]), [
            ("vk_Status", 0, "OK"),
            ("vk_PreisAngesetzt", 0, 20 * miete_2030),
            ("vk_BuchwertGeb", 0, 320_000),
            ("vk_BuchwertGesamt", 0, 520_000),
            ("vk_ErloesGeb", 0, netto_2030 * 320_000 / 520_000),
            ("vk_Gewinn", 0, netto_2030 - 520_000),
            ("prg_Aktiv", pz(0, 2030), 1),
            ("prg_Buchwert", pz(0, 2030), 320_000),
            ("prg_Aktiv", pz(0, 2031), 0),
            ("prg_Ergebnis", pz(0, 2031), 0),
        ]),
        ("Etappe 4: Verkauf eines von zwei Objekten",
         Modell([obj, kurz], verkaeufe=[dataclasses.replace(vk_2030, objekt_id="OBJ-002")]), [
            ("prg_Aktiv", pz(0, 2031), 1),       # erstes Objekt läuft weiter
            ("prg_AfA", pz(0, 2031), 20_000),
            ("prg_Aktiv", pz(1, 2031), 0),
            ("vk_BuchwertGeb", 0, 0),            # OBJ-002 ist 2029 abgeschrieben
            ("vk_Gewinn", 0, netto_2030 - 200_000),
        ]),
        ("Etappe 4: ObjektID unbekannt", vk_status(dataclasses.replace(vk_2027, objekt_id="X")),
         [("vk_Status", 0, "ObjektID unbekannt"), ("vk_Gewinn", 0, None),
          ("prg_Aktiv", pz(0, 2046), 1)]),
        ("Etappe 4: Objekt mit Fehlerstatus", vk_status(vk_2027, [ohne_miete]),
         [("vk_Status", 0, "Objekt nicht OK")]),
        ("Etappe 4: Verkauf doppelt", Modell([obj], verkaeufe=[vk_2027, vk_2030]),
         [("vk_Status", 0, "Verkauf doppelt"), ("vk_Status", 1, "Verkauf doppelt"),
          ("prg_Aktiv", pz(0, 2046), 1)]),  # ungültige Verkäufe schalten nichts ab
        ("Etappe 4: 6b-Angabe fehlt", vk_status(dataclasses.replace(vk_2027, nutzung_6b=None)),
         [("vk_Status", 0, "Pflichtfeld fehlt")]),
        ("Etappe 4: Preis und Faktor zugleich",
         vk_status(dataclasses.replace(vk_2027, faktor=20)),
         [("vk_Status", 0, "Preis oder Faktor angeben")]),
        ("Etappe 4: weder Preis noch Faktor", vk_status(dataclasses.replace(vk_2027, preis=None)),
         [("vk_Status", 0, "Preis oder Faktor angeben")]),
        ("Etappe 4: Verkaufsjahr außerhalb", vk_status(dataclasses.replace(vk_2027, jahr=2050)),
         [("vk_Status", 0, "Verkaufsjahr außerhalb Prognose")]),
        ("Etappe 4: Verkehrswert ohne Gebäudeanteil",
         vk_status(dataclasses.replace(vk_2027, aufteilung="Verkehrswert"), [ohne_quote]),
         [("vk_Status", 0, "Gebäudeanteil fehlt")]),
        ("Etappe 4: Standardaufteilung Verkehrswert vom Parameterblatt",
         Modell([vk_obj], {"par_Aufteilung": "Verkehrswert"},
                [dataclasses.replace(vk_2027, aufteilung=None)]),
         [("vk_ErloesGeb", 0, 700_000), ("vk_Gewinn", 0, 720_000)]),
        ("Etappe 5: Rücklage voll, Abnahmefall 720.000",
         Modell([vk_obj], verkaeufe=[vk_2027]), [
            ("rl_ID", 0, "OBJ-001"),
            ("rl_Jahr", 0, 2027),
            ("rl_Kaufjahr", 0, 2007),
            ("rl_Vorbesitz", 0, 20),
            ("rl_Status", 0, "Rücklage gebildet"),
            ("rl_RuecklageID", 0, "R-OBJ-001"),
            ("rl_Gewinn", 0, 720_000),
            ("rl_BetragGeb", 0, gewinn_geb_bw),
            ("rl_BetragGuB", 0, gewinn_gub_bw),
            ("rl_Ruecklage", 0, 720_000),
            ("rl_Steuerpflichtig", 0, 0),
            ("rl_Fristjahr", 0, 2031),
            ("rl_UebertragGeb", 0, 0),
            ("rl_Rest", 0, 720_000),
            ("rl_Zuschlag", 0, zuschlag),
            ("rl_ID", 1, None),                  # leere Zeile bleibt leer
            ("rl_Status", 1, None),
            ("rl_Fristjahr", 1, None),
            ("rlj_Jahr", rj(2027), 2027),
            ("rlj_Gewinn", rj(2027), 720_000),
            ("rlj_Bildung", rj(2027), 720_000),
            ("rlj_Stand", rj(2027), 720_000),
            ("rlj_Steuerpflichtig", rj(2027), 0),
            ("rlj_Steuer", rj(2027), 0),         # Steuer im Verkaufsjahr null
            ("rlj_Stand", rj(2030), 720_000),
            ("rlj_Aufloesung", rj(2030), 0),
            ("rlj_Aufloesung", rj(2031), 720_000),
            ("rlj_Zuschlag", rj(2031), zuschlag),
            ("rlj_Stand", rj(2031), 0),
            ("rlj_Steuerpflichtig", rj(2031), 720_000 + zuschlag),
            ("rlj_Steuer", rj(2031), (720_000 + zuschlag) * 0.30),
            ("rlj_Stand", rj(2046), 0),
        ]),
        ("Etappe 5: Rücklage nach Verkehrswert getrennt",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, aufteilung="Verkehrswert")]), [
            ("rl_BetragGeb", 0, 220_000),
            ("rl_BetragGuB", 0, 500_000),
            ("rl_Ruecklage", 0, 720_000),
        ]),
        ("Etappe 5: ohne 6b sofort steuerpflichtig", Modell([obj], verkaeufe=[vk_2030]), [
            ("rl_Status", 0, "6b nicht gewählt"),
            ("rl_RuecklageID", 0, None),
            ("rl_Ruecklage", 0, 0),
            ("rl_Fristjahr", 0, None),
            ("rl_Zuschlag", 0, 0),
            ("rl_Steuerpflichtig", 0, gewinn_2030),
            ("rlj_Steuerpflichtig", rj(2030), gewinn_2030),
            ("rlj_Steuer", rj(2030), gewinn_2030 * 0.30),
            ("rlj_Stand", rj(2030), 0),
            ("rlj_Steuer", rj(2029), 0),
        ]),
        ("Etappe 5: Vorbesitzzeit 5 Jahre zu kurz",
         Modell([dataclasses.replace(vk_obj, kaufjahr=2022)], verkaeufe=[vk_2027]), [
            ("rl_Vorbesitz", 0, 5),
            ("rl_Status", 0, "Vorbesitzzeit zu kurz"),
            ("rl_Ruecklage", 0, 0),
            ("rl_Steuerpflichtig", 0, 720_000),
            ("rlj_Steuer", rj(2027), 216_000),
        ]),
        ("Etappe 5: Vorbesitzzeit genau 6 Jahre",
         Modell([dataclasses.replace(vk_obj, kaufjahr=2021)], verkaeufe=[vk_2027]), [
            ("rl_Status", 0, "Rücklage gebildet"),
            ("rl_Ruecklage", 0, 720_000),
        ]),
        ("Etappe 5: Verlust, keine Rücklage",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, preis=600_000)]), [
            ("rl_Status", 0, "kein Gewinn"),
            ("rl_Ruecklage", 0, 0),
            ("rl_Steuerpflichtig", 0, -80_000),
            ("rlj_Steuer", rj(2027), -24_000),
        ]),
        ("Etappe 5: Verlust Gebäude, Gewinn G+B",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, preis=700_000,
                                                         aufteilung="Verkehrswert")]), [
            ("rl_Status", 0, "Rücklage gebildet"),
            ("rl_BetragGeb", 0, 0),              # Gebäudeverlust 130.000 wirkt sofort
            ("rl_BetragGuB", 0, 150_000),
            ("rl_Steuerpflichtig", 0, -130_000),
            ("rl_Zuschlag", 0, 150_000 * 0.24),
        ]),
        ("Etappe 5: Fristjahr nach Prognoseende",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, jahr=2044)]), [
            ("rl_Ruecklage", 0, 1_060_000),      # Buchwert 2044: 140.000 + 200.000
            ("rl_Fristjahr", 0, 2048),
            ("rlj_Stand", rj(2046), 1_060_000),
            ("rlj_Aufloesung", rj(2046), 0),
            ("rlj_Steuer", rj(2044), 0),
        ]),
        ("Etappe 5: Frist 6 Jahre, Zuschlag 5 %, Steuersatz 25 %",
         Modell([vk_obj], {"par_6bFrist": 6, "par_6bZuschlag": 0.05, "par_Steuersatz": 0.25},
                [vk_2027]), [
            ("rl_Fristjahr", 0, 2033),
            ("rl_Zuschlag", 0, 216_000),
            ("rlj_Stand", rj(2032), 720_000),
            ("rlj_Steuerpflichtig", rj(2033), 936_000),
            ("rlj_Steuer", rj(2033), 234_000),
        ]),
        ("Etappe 5: zwei Verkäufe im selben Jahr",
         Modell([vk_obj, kurz], verkaeufe=[
             vk_2027, Verkauf("OBJ-002", 2027, preis=500_000, nutzung_6b="nein")]), [
            ("rl_Status", 1, "6b nicht gewählt"),
            ("rl_Gewinn", 1, 270_000),           # Buchwert 30.000 + G+B 200.000
            ("rlj_Gewinn", rj(2027), 990_000),
            ("rlj_Bildung", rj(2027), 720_000),
            ("rlj_Steuerpflichtig", rj(2027), 270_000),
            ("rlj_Steuer", rj(2027), 81_000),
        ]),
        ("Etappe 5: ungültiger Verkauf bildet nichts",
         vk_status(dataclasses.replace(vk_2027, objekt_id="X")), [
            ("rl_ID", 0, None),
            ("rl_Status", 0, None),
            ("rlj_Gewinn", rj(2027), 0),
            ("rlj_Stand", rj(2027), 0),
        ]),
        ("Etappe 6: Abnahmefall Übertrag, AfA-Basis 180.000", reinvest(neu()), [
            ("neu_Status", 0, "OK"),
            ("neu_Status", 1, None),             # leere Zeile bleibt leer
            ("neu_AKGesamt", 0, 900_000),
            ("neu_AKGeb", 0, 900_000),
            ("neu_UebGeb", 0, gewinn_geb_bw),
            ("neu_UebGuBGuB", 0, 0),
            ("neu_UebGuBGeb", 0, gewinn_gub_bw),  # G+B-Gewinn darf aufs Gebäude
            ("neu_Uebertrag", 0, 720_000),
            ("neu_AfABasis", 0, 180_000),
            ("vk_Gewinn", 0, 720_000),           # kein Zirkelbezug über die Prognose
            ("rl_UebertragGeb", 0, gewinn_geb_bw),
            ("rl_UebertragGuB", 0, gewinn_gub_bw),
            ("rl_Rest", 0, 0),
            ("rl_Zuschlag", 0, 0),
            ("rlj_Stand", rj(2027), 720_000),
            ("rlj_Uebertrag", rj(2028), 720_000),
            ("rlj_Stand", rj(2028), 0),
            ("rlj_Aufloesung", rj(2031), 0),
            ("rlj_Steuer", rj(2031), 0),
            ("prg_ID", nz(0, 2027), "NEU-001"),
            ("prg_Aktiv", nz(0, 2028), 0),       # Kauf zum Jahresende
            ("prg_AfA", nz(0, 2028), 0),
            ("prg_Buchwert", nz(0, 2027), 0),
            ("prg_Buchwert", nz(0, 2028), 180_000),
            ("prg_Aktiv", nz(0, 2029), 1),
            ("prg_AfA", nz(0, 2029), 5_400),     # 3 % von 180.000, nicht von 900.000
            ("prg_Buchwert", nz(0, 2029), 174_600),
            ("prg_Miete", nz(0, 2029), 45_000),
            ("prg_Miete", nz(0, 2030), 45_900),
            ("prg_Erhaltung", nz(0, 2029), 9_000),
            ("prg_Ergebnis", nz(0, 2029), 45_000 - 9_000 - 5_400),
            ("prg_Buchwert", nz(0, 2046), 180_000 - 18 * 5_400),
            ("prg_Aktiv", pz(0, 2028), 0),       # Altobjekt bleibt verkauft
        ]),
        ("Etappe 6: G+B-Rücklage zuerst auf G+B",
         reinvest(neu(kaufpreis=1_000_000, anteil_gub=0.1)), [
            ("neu_AKGeb", 0, 900_000),
            ("neu_AKGuB", 0, 100_000),
            ("neu_UebGuBGuB", 0, 100_000),
            ("neu_UebGuBGeb", 0, gewinn_gub_bw - 100_000),
            ("neu_AfABasis", 0, 280_000),
            ("neu_BuchwertGuB", 0, 0),
            ("rl_Rest", 0, 0),
        ]),
        ("Etappe 6: Gebäude-Rücklage nicht auf G+B",
         reinvest(neu(kaufpreis=1_000_000, anteil_gub=1.0)), [
            ("neu_AKGeb", 0, 0),
            ("neu_UebGeb", 0, 0),
            ("neu_UebGuBGuB", 0, gewinn_gub_bw),
            ("neu_BuchwertGuB", 0, 1_000_000 - gewinn_gub_bw),
            ("rl_Rest", 0, gewinn_geb_bw),
            ("rl_Zuschlag", 0, gewinn_geb_bw * 0.24),
            ("rlj_Aufloesung", rj(2031), gewinn_geb_bw),
        ]),
        ("Etappe 6: Neuobjekt zu klein, Rest wird aufgelöst",
         reinvest(neu(kaufjahr=2029, kaufpreis=400_000, anteil_gub=0.25)), [
            ("neu_UebGeb", 0, 300_000),
            ("neu_UebGuBGuB", 0, 100_000),
            ("neu_UebGuBGeb", 0, 0),
            ("neu_AfABasis", 0, 0),
            ("rl_Rest", 0, rest_teil),
            ("rl_Zuschlag", 0, rest_teil * 0.24),
            ("rlj_Uebertrag", rj(2029), 400_000),
            ("rlj_Stand", rj(2029), rest_teil),
            ("rlj_Aufloesung", rj(2031), rest_teil),
            ("rlj_Steuer", rj(2031), rest_teil * 1.24 * 0.30),
            ("rlj_Stand", rj(2031), 0),
            ("prg_AfA", nz(0, 2030), 0),
        ]),
        ("Etappe 6: zwei Neuobjekte aus einer Rücklage",
         reinvest(neu(kaufjahr=2029, kaufpreis=400_000, anteil_gub=0.25),
                  neu("NEU-002", 2030, 1_000_000, 0.2)), [
            ("neu_UebGeb", 1, gewinn_geb_bw - 300_000),
            ("neu_UebGuBGuB", 1, gewinn_gub_bw - 100_000),
            ("neu_UebGuBGeb", 1, 0),
            ("neu_AfABasis", 1, basis_zweit),
            ("neu_BuchwertGuB", 1, 200_000 - (gewinn_gub_bw - 100_000)),
            ("rl_UebertragGeb", 0, gewinn_geb_bw),
            ("rl_Rest", 0, 0),
            ("rlj_Uebertrag", rj(2030), rest_teil),
            ("rlj_Stand", rj(2030), 0),
            ("rlj_Steuer", rj(2031), 0),
            ("prg_ID", nz(1, 2027), "NEU-002"),
            ("prg_AfA", nz(1, 2031), basis_zweit * 0.03),
        ]),
        ("Etappe 6: Neuobjekt ohne Rücklage, Nebenkosten aus GrESt",
         Modell([obj], neuobjekte=[neu(kaufjahr=2030, kaufpreis=500_000, anteil_gub=0.2,
                                       nebenkosten=None, quelle=None)]), [
            ("neu_Status", 0, "OK"),
            ("neu_AKGesamt", 0, 525_000),
            ("neu_AKGeb", 0, 420_000),
            ("neu_AKGuB", 0, 105_000),
            ("neu_Uebertrag", 0, 0),
            ("neu_AfABasis", 0, 420_000),
            ("prg_Buchwert", nz(0, 2029), 0),
            ("prg_Buchwert", nz(0, 2030), 420_000),
            ("prg_Miete", nz(0, 2030), 0),
            ("prg_Miete", nz(0, 2031), 25_000),
            ("prg_Miete", nz(0, 2032), 25_500),
            ("prg_Erhaltung", nz(0, 2031), 5_000),
            ("prg_AfA", nz(0, 2031), 12_600),
            ("prg_Buchwert", nz(0, 2031), 407_400),
            ("prg_Ergebnis", nz(0, 2031), 25_000 - 5_000 - 12_600),
        ]),
        ("Etappe 6: Kauf nach Fristjahr", reinvest(neu(kaufjahr=2032)), [
            ("neu_Status", 0, "Kauf nach Fristjahr"),
            ("neu_AfABasis", 0, None),
            ("rl_Rest", 0, 720_000),
            ("rlj_Aufloesung", rj(2031), 720_000),
            ("prg_Aktiv", nz(0, 2033), 0),
            ("prg_Buchwert", nz(0, 2033), 0),
        ]),
        ("Etappe 6: Kauf vor Verkauf",
         reinvest(neu(kaufjahr=2029), verkauf=dataclasses.replace(vk_2030, nutzung_6b="ja"),
                  objekte=(obj,)),
         [("neu_Status", 0, "Kauf vor Verkauf"), ("rl_UebertragGeb", 0, 0)]),
        ("Etappe 6: Rücklage unbekannt", reinvest(neu(quelle="R-X")),
         [("neu_Status", 0, "Rücklage unbekannt")]),
        ("Etappe 6: Rücklage eines Verkaufs ohne 6b",
         reinvest(neu(), verkauf=dataclasses.replace(vk_2027, nutzung_6b="nein")),
         [("neu_Status", 0, "Rücklage unbekannt")]),
        ("Etappe 6: NeuID wie Bestandsobjekt", reinvest(neu(neu_id="OBJ-001")),
         [("neu_Status", 0, "NeuID doppelt"), ("rl_Rest", 0, 720_000)]),
        ("Etappe 6: Pflichtfeld fehlt",
         reinvest(dataclasses.replace(neu(), mietrendite=None)),
         [("neu_Status", 0, "Pflichtfeld fehlt")]),
        ("Etappe 6: Kaufjahr außerhalb", reinvest(neu(kaufjahr=2050, quelle=None)),
         [("neu_Status", 0, "Kaufjahr außerhalb Prognose")]),
        ("Etappe 7: ein Objekt ohne Verkauf", Modell([obj]), [
            ("liq_Jahr", rj(2027), 2027),
            ("liq_Miete", rj(2027), 61_200),
            ("liq_Erhaltung", rj(2027), 8_200),
            ("liq_AfA", rj(2027), 20_000),
            ("liq_Ergebnis", rj(2027), 33_000),
            ("liq_Ueberschuss", rj(2027), 53_000),  # AfA fließt nicht ab
            ("liq_Erloes", rj(2027), 0),
            ("liq_SteuerLaufend", rj(2027), 9_900),
            ("liq_SteuerVerkauf", rj(2027), 0),
            ("liq_Steuer", rj(2027), 9_900),
            ("liq_Reinvest", rj(2027), 0),
            ("liq_Mittelzufluss", rj(2027), fluss_2027),
            ("liq_MittelzuflussKum", rj(2028), fluss_2027 + fluss_2028),
            ("liq_Jahr", summe, "Summe"),
            ("liq_Ergebnis", summe, sum(ergebnisse)),
            ("liq_Steuer", summe, sum(ergebnisse) * 0.30),
            ("liq_MittelzuflussKum", summe, None),  # kumulierte Werte ohne Summe
            ("prg_Bestand", pz(0, 2027), 1),
            ("prg_BuchwertGuB", pz(0, 2027), 200_000),
            ("prg_BuchwertGesamt", pz(0, 2027), 580_000),
            ("prg_Verkehrswert", pz(0, 2027), 1_428_000),
            ("prg_StilleReserven", pz(0, 2027), 848_000),
            ("prg_Bestand", pz(1, 2027), 0),     # leerer Block zählt nicht
            ("prg_Verkehrswert", pz(1, 2027), 0),
            ("aw_Ergebnis", rj(2027), 33_000),
            ("aw_SteuerpflichtigVerkauf", rj(2027), 0),
            ("aw_GuV", rj(2027), 33_000),
            ("aw_Steuer", rj(2027), 9_900),
            ("aw_NachSteuer", rj(2027), 23_100),
            ("aw_Buchwert", rj(2027), 580_000),
            ("aw_Verkehrswert", rj(2027), 1_428_000),
            ("aw_StilleReserven", rj(2027), 848_000),
            ("aw_Buchwert", rj(2046), 200_000),
            ("aw_Verkehrswert", rj(2046), 1_400_000 * 1.02 ** 20),
            ("aw_SteuerKum", rj(2046), sum(ergebnisse) * 0.30),
            ("aw_Steuer", summe, sum(ergebnisse) * 0.30),
            ("aw_Ruecklage", rj(2027), 0),
            ("aw_MittelzuflussKum", rj(2028), fluss_2027 + fluss_2028),
        ]),
        ("Etappe 7: Summe über zwei Objekte", Modell([obj, kurz]), [
            ("liq_Miete", rj(2030), 2 * miete_2030),
            ("liq_AfA", rj(2029), 30_000),
            ("liq_AfA", rj(2030), 20_000),       # OBJ-002 ist abgeschrieben
            ("liq_Ergebnis", rj(2030), ergebnis_2030 + miete_2030 - erh_2030),
            ("aw_Buchwert", rj(2030), 720_000),
            ("aw_Verkehrswert", rj(2030), 2 * 1_400_000 * 1.02 ** 4),
        ]),
        ("Etappe 7: Verkauf ohne 6b, Steuer = Gewinn × Satz",
         Modell([obj], verkaeufe=[vk_2030]), [
            ("liq_Erloes", rj(2030), netto_2030),
            ("liq_BuchwertRueckfluss", rj(2030), 520_000),
            ("liq_Gewinn", rj(2030), gewinn_2030),
            ("liq_SteuerVerkauf", rj(2030), gewinn_2030 * 0.30),
            ("liq_Steuer", rj(2030), steuer_2030),
            ("liq_Mittelzufluss", rj(2030), miete_2030 - erh_2030 + netto_2030 - steuer_2030),
            ("liq_Miete", rj(2031), 0),
            ("liq_Mittelzufluss", rj(2031), 0),
            ("prg_Bestand", pz(0, 2029), 1),
            ("prg_Bestand", pz(0, 2030), 0),     # Verkauf zum Jahresende
            ("aw_Buchwert", rj(2029), 540_000),
            ("aw_Buchwert", rj(2030), 0),
            ("aw_Verkehrswert", rj(2030), 0),
            ("aw_GuV", rj(2030), ergebnis_2030 + gewinn_2030),
            ("aw_Steuer", rj(2030), steuer_2030),
        ]),
        ("Etappe 7: Reinvestition, Abnahmefall", reinvest(neu()), [
            ("liq_Erloes", rj(2027), 1_400_000),
            ("liq_SteuerVerkauf", rj(2027), 0),
            ("liq_Steuer", rj(2027), 9_900),
            ("liq_Mittelzufluss", rj(2027), 53_000 + 1_400_000 - 9_900),
            ("liq_Reinvest", rj(2028), 900_000),
            ("liq_Miete", rj(2028), 0),
            ("liq_Mittelzufluss", rj(2028), -900_000),
            ("liq_MittelzuflussKum", rj(2028), 53_000 + 1_400_000 - 9_900 - 900_000),
            ("liq_Ergebnis", rj(2029), 30_600),
            ("liq_Steuer", rj(2029), 9_180),
            ("liq_Mittelzufluss", rj(2029), 45_000 - 9_000 - 9_180),
            ("liq_SteuerVerkauf", rj(2031), 0),
            ("prg_Bestand", nz(0, 2027), 0),
            ("prg_Bestand", nz(0, 2028), 1),
            ("prg_BuchwertGesamt", nz(0, 2028), 180_000),
            ("aw_Buchwert", rj(2027), 0),        # Altobjekt verkauft, Neuobjekt noch nicht da
            ("aw_Ruecklage", rj(2027), 720_000),
            ("aw_Ruecklage", rj(2028), 0),
            ("aw_Buchwert", rj(2028), 180_000),
            ("aw_Verkehrswert", rj(2028), 900_000),
            ("aw_StilleReserven", rj(2028), 720_000),  # = übertragene Rücklage
            ("aw_Verkehrswert", rj(2030), 900_000 * 1.02 ** 2),
            ("aw_Buchwert", rj(2029), 174_600),
        ]),
        ("Etappe 7: Neuobjekt mit G+B im Buchwert",
         reinvest(neu(kaufpreis=1_000_000, anteil_gub=0.1)), [
            ("prg_BuchwertGuB", nz(0, 2028), 0),
            ("aw_Buchwert", rj(2028), 280_000),
            ("aw_StilleReserven", rj(2028), 720_000),
        ]),
        ("Etappe 7: Neuobjekt ohne Rücklage", Modell([obj], neuobjekte=[
            neu(kaufjahr=2030, kaufpreis=500_000, anteil_gub=0.2, nebenkosten=None,
                quelle=None)]), [
            ("liq_Reinvest", rj(2030), 525_000),
            ("prg_BuchwertGuB", nz(0, 2030), 105_000),
            ("prg_BuchwertGesamt", nz(0, 2030), 525_000),
            ("prg_StilleReserven", nz(0, 2030), -25_000),  # Nebenkosten nicht im Verkehrswert
            ("aw_Buchwert", rj(2030), 520_000 + 525_000),
        ]),
        ("Etappe 7: ungültiges Neuobjekt fließt nicht ab", reinvest(neu(kaufjahr=2032)), [
            ("liq_Reinvest", rj(2032), 0),
            ("liq_SteuerVerkauf", rj(2031), (720_000 + zuschlag) * 0.30),
            ("aw_SteuerpflichtigVerkauf", rj(2031), 720_000 + zuschlag),
            ("aw_Verkehrswert", rj(2032), 0),
        ]),
        ("Etappe 7: ohne Verkehrswert keine stille Reserve",
         Modell([dataclasses.replace(obj, verkehrswert=None)]), [
            ("prg_Verkehrswert", pz(0, 2027), 580_000),
            ("aw_StilleReserven", rj(2027), 0),
        ]),
        ("Etappe 7: Verlust mindert die Steuer im selben Jahr",
         Modell([vk_obj], verkaeufe=[dataclasses.replace(vk_2027, preis=600_000)]), [
            ("liq_Erloes", rj(2027), 600_000),
            ("liq_Steuer", rj(2027), (33_000 - 80_000) * 0.30),
            ("aw_GuV", rj(2027), 33_000 - 80_000),
        ]),
        ("Etappe 8: Alternativanlage und Vermögen, ein Objekt", Modell([obj]), [
            ("aw_Zins", rj(2027), 0),            # Anfangsbestand null
            ("aw_Anlage", rj(2027), fluss_2027),
            ("aw_Zins", rj(2028), fluss_2027 * 0.04),
            ("aw_SteuerZins", rj(2028), fluss_2027 * 0.04 * 0.30),
            ("aw_Anlage", rj(2028), fluss_2027 * 1.028 + fluss_2028),
            ("aw_LatenteSteuer", rj(2027), 848_000 * 0.30),
            ("aw_Vermoegen", rj(2027), 1_428_000 + fluss_2027 - 848_000 * 0.30),
            ("aw_Anlage", summe, None),          # Bestand ohne Summe
            ("vg_Aktuell", vg("verkehrswert"), 1_400_000 * 1.02 ** 20),
            ("vg_Aktuell", vg("afa"), 20 * 20_000),
            ("vg_Aktuell", vg("ergebnis"), sum(ergebnisse)),
            ("vg_A", vg("vermoegen"), None),     # nichts gespeichert
            ("vg_Differenz", vg("vermoegen"), None),
        ]),
        ("Etappe 8: Szenario A, 6b-Kette", szenario_a, [
            ("rl_Status", 0, "Rücklage gebildet"),
            ("neu_Status", 0, "OK"),
            ("aw_Anlage", rj(2028), 53_000 + 1_400_000 - 9_900 - 900_000),
            ("aw_Zins", rj(2029), 0),
            ("aw_Ruecklage", rj(2027), 720_000),
            ("aw_LatenteSteuer", rj(2027), 720_000 * 0.30),  # Rücklage ist latent steuerpflichtig
            ("vg_Aktuell", vg("verkehrswert"), vw_neu),
            ("vg_Aktuell", vg("buchwert"), bw_neu),
            ("vg_Aktuell", vg("anlage"), anlage_a),
            ("vg_Aktuell", vg("latente_steuer"), latent_a),
            ("vg_Aktuell", vg("vermoegen"), vermoegen_a),
            ("vg_Aktuell", vg("ruecklage"), 0),
            ("vg_Aktuell", vg("afa"), 20_000 + 18 * 5_400),  # AfA nur auf 180.000
            ("vg_Aktuell", vg("reinvest"), 900_000),
            ("vg_Aktuell", vg("zins"), 0),
        ]),
        ("Etappe 8: Szenario B, sofort versteuern", szenario_b, [
            ("rl_Status", 0, RUECKLAGE_SZENARIO_B),
            ("rl_Ruecklage", 0, 0),
            ("rl_RuecklageID", 0, None),
            ("rlj_Steuer", rj(2027), 216_000),
            ("neu_Status", 0, NEU_SZENARIO_B),
            ("liq_Reinvest", rj(2028), 0),
            ("liq_Steuer", rj(2027), steuer_b_2027),
            ("liq_Mittelzufluss", rj(2027), anlage_b),
            ("liq_Mittelzufluss", rj(2029), 0),
            ("prg_Aktiv", nz(0, 2029), 0),       # kein Neuobjekt
            ("vg_Aktuell", vg("verkehrswert"), 0),
            ("vg_Aktuell", vg("latente_steuer"), 0),
            ("vg_Aktuell", vg("anlage"), anlage_b),
            ("vg_Aktuell", vg("vermoegen"), anlage_b),
            ("vg_Aktuell", vg("steuer"), steuer_b_2027),
            ("vg_Aktuell", vg("steuer_gesamt"), steuer_b_2027),
            ("vg_Aktuell", vg("afa"), 20_000),
        ]),
        ("Etappe 8: Szenario B mit Alternativrendite 4 %",
         dataclasses.replace(szenario_b, parameter={"par_Szenario": "B"}), [
            ("aw_Zins", rj(2028), anlage_b * 0.04),
            ("vg_Aktuell", vg("anlage"), anlage_b * 1.028 ** 19),
            ("vg_Aktuell", vg("zins"), zins_b),
            ("vg_Aktuell", vg("steuer_zins"), zins_b * 0.30),
            ("vg_Aktuell", vg("steuer_gesamt"), steuer_b_2027 + zins_b * 0.30),
            ("vg_Aktuell", vg("vermoegen"), anlage_b * 1.028 ** 19),
        ]),
        ("Etappe 8: Neuobjekt ohne Rücklage bleibt in Szenario B", Modell([obj], {
            "par_Szenario": "B"}, neuobjekte=[neu(kaufjahr=2030, kaufpreis=500_000,
                                                  nebenkosten=None, quelle=None)]), [
            ("neu_Status", 0, "OK"),
            ("liq_Reinvest", rj(2030), 525_000),
        ]),
    ] + [
        (f"Etappe 1: Steuerwelt {welt}", Modell([obj], {"par_Steuerwelt": welt}),
         [("par_StatusSteuerwelt", 0, "nicht im MVP – Ergebnisse gelten nur für GmbH")])
        for welt in STEUERWELTEN[1:]
    ]


def pruefe(fall: str, wb, pruefungen) -> int:
    """Sollwerte eines Falls prüfen, Ergebnis ausgeben, Anzahl Abweichungen zurückgeben."""
    fehler = 0
    for name, zeile, soll in pruefungen:
        ist = wert(wb, name, zeile)
        if ist == "":
            ist = None
        if gleich(ist, soll):
            print(f"OK      {fall}: {name}[{zeile}] = {ist!r}")
        else:
            fehler += 1
            print(f"FEHLER  {fall}: {name}[{zeile}] Soll {soll!r}, Ist {ist!r}")
    return fehler


def gespeicherter_vergleich(laeufe: dict, ordner: Path) -> int:
    """Etappe 8: Läufe A und B in die Vergleichsspalten übernehmen, Differenz prüfen.

    Spielt nach, was ab Etappe 9 das Makro tut: Kennzahlen des aktiven Szenarios
    als Werte in Spalte A bzw. B speichern.
    """
    werte = {sz: {k.key: wert(wb, "vg_Aktuell", vg(k.key)) for k in VERGLEICH_KENNZAHLEN}
             for sz, (modell, wb) in laeufe.items()}
    modell = dataclasses.replace(laeufe["B"][0], vergleich=werte)
    wb = durchrechnen(modell, ordner)
    pruefungen = []
    for k in VERGLEICH_KENNZAHLEN:
        a, b = werte["A"][k.key], werte["B"][k.key]
        pruefungen += [("vg_A", vg(k.key), a), ("vg_B", vg(k.key), b),
                       ("vg_Differenz", vg(k.key), a - b)]
    fehler = pruefe("Etappe 8: gespeicherter Vergleich A − B", wb, pruefungen)
    if werte["A"]["vermoegen"] <= werte["B"]["vermoegen"]:
        fehler += 1
        print("FEHLER  Etappe 8: ohne Alternativrendite muss A beim Endvermögen vorn liegen")
    else:
        print("OK      Etappe 8: ohne Alternativrendite liegt A beim Endvermögen vorn")
    return fehler


def main() -> int:
    """Alle Fälle prüfen; optional nur Fälle, deren Name den ersten Aufrufparameter enthält."""
    filter_ = sys.argv[1] if len(sys.argv) > 1 else ""
    fehler = 0
    laeufe = {}
    with tempfile.TemporaryDirectory() as tmp:
        for i, (fall, modell, pruefungen) in enumerate(faelle()):
            if filter_ not in fall:
                continue
            wb = durchrechnen(modell, Path(tmp) / f"fall{i}")
            fehler += pruefe(fall, wb, pruefungen)
            if fall.startswith("Etappe 8: Szenario A,"):
                laeufe["A"] = (modell, wb)
            elif fall.startswith("Etappe 8: Szenario B,"):
                laeufe["B"] = (modell, wb)
        if len(laeufe) == 2:
            fehler += gespeicherter_vergleich(laeufe, Path(tmp) / "vergleich")
    print(f"\n{fehler} Abweichung(en)" if fehler else "\nAlle Prüfungen bestanden.")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main())
