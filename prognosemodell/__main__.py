"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten]"""

import argparse
from pathlib import Path

from .mappe import erstelle_mappe
from .modelle import Modell
from .testdaten import testmodell


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ausgabe", default="ausgabe/Prognosemodell_VV.xlsx")
    ap.add_argument("--ohne-testdaten", action="store_true",
                    help="leeres Objektblatt statt Testobjekt")
    args = ap.parse_args()

    modell = Modell() if args.ohne_testdaten else testmodell()
    ziel = Path(args.ausgabe)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    erstelle_mappe(modell).save(ziel)
    print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
