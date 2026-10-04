"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten] [--ohne-makros]
                                             [--stammdaten DATEI --kostenstellen DATEI]

Mit --stammdaten und --kostenstellen kommen die Objekte aus den beiden
Eingabedateien (Vorlagen: python -m prognosemodell.vorlagen), sonst das Testobjekt.

Standard ist die .xlsm mit eingebetteter VBA-Steuerung; dafür braucht der Bau
LibreOffice (siehe makros.py). Mit --ohne-makros entsteht eine reine .xlsx.
"""

import argparse
import sys
from pathlib import Path

from .einlesen import EinleseFehler, lies_modell
from .mappe import erstelle_mappe
from .modelle import Modell
from .testdaten import testmodell

AUSGABE = "ausgabe/Prognosemodell_VV"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ausgabe", help=f"Zieldatei, Standard {AUSGABE}.xlsm bzw. .xlsx")
    ap.add_argument("--ohne-testdaten", action="store_true",
                    help="leeres Objektblatt statt Testobjekt")
    ap.add_argument("--ohne-makros", action="store_true",
                    help="reine .xlsx ohne VBA, kein LibreOffice nötig")
    ap.add_argument("--stammdaten", help="Stammdatendatei (Objekte, Kontenzuordnung)")
    ap.add_argument("--kostenstellen", help="Kostenstellendatei, ein Blatt je Objekt")
    args = ap.parse_args()

    if bool(args.stammdaten) != bool(args.kostenstellen):
        ap.error("--stammdaten und --kostenstellen nur zusammen angeben.")
    if args.stammdaten:
        try:
            ergebnis = lies_modell(args.stammdaten, args.kostenstellen)
        except EinleseFehler as e:
            print("Einlesen abgebrochen:", *e.meldungen, sep="\n  ", file=sys.stderr)
            sys.exit(1)
        for hinweis in ergebnis.hinweise:
            print(f"Hinweis: {hinweis}")
        modell = ergebnis.modell
        print(f"eingelesen: {len(modell.objekte)} Objekte, Basisjahr "
              f"{modell.parameter.get('par_Basisjahr')}")
    else:
        modell = Modell() if args.ohne_testdaten else testmodell()
    endung = ".xlsx" if args.ohne_makros else ".xlsm"
    ziel = Path(args.ausgabe or AUSGABE + endung)
    if ziel.suffix != endung:
        ap.error(f"Die Zieldatei muss auf {endung} enden.")
    ziel.parent.mkdir(parents=True, exist_ok=True)
    wb = erstelle_mappe(modell)
    if args.ohne_makros:
        wb.save(ziel)
    else:
        from .makros import speichere_mit_makros
        speichere_mit_makros(wb, ziel)
    print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
