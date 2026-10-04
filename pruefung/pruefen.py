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
