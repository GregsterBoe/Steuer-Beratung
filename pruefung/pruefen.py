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
from prognosemodell.modelle import STEUERWELTEN, Modell, Verkauf
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
    ] + [
        (f"Etappe 1: Steuerwelt {welt}", Modell([obj], {"par_Steuerwelt": welt}),
         [("par_StatusSteuerwelt", 0, "nicht im MVP – Ergebnisse gelten nur für GmbH")])
        for welt in STEUERWELTEN[1:]
    ]


def main() -> int:
    fehler = 0
    with tempfile.TemporaryDirectory() as tmp:
        for i, (fall, modell, pruefungen) in enumerate(faelle()):
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
