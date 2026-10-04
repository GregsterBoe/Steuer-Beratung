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
from prognosemodell.modelle import STEUERWELTEN, Modell
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
    erh_2030 = 8_000 * 1.025 ** 4
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
