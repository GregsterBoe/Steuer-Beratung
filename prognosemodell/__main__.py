"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten] [--ohne-makros]

Standard ist die .xlsm mit eingebetteter VBA-Steuerung; dafür braucht der Bau
LibreOffice (siehe makros.py). Mit --ohne-makros entsteht eine reine .xlsx.
"""

import argparse
from pathlib import Path

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
    args = ap.parse_args()

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
