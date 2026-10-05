"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten]
[--kostenstellen PFAD]"""

import argparse
from pathlib import Path

from .einlesen import lese_kostenstellen, zusammenfuehren
from .mappe import erstelle_mappe
from .modelle import PARAMETER, Modell
from .testdaten import testmodell


def _basisjahr() -> int:
    return next(p.wert for p in PARAMETER if p.name == "par_Basisjahr")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--ausgabe", default="ausgabe/Prognosemodell_VV.xlsx")
    ap.add_argument("--ohne-testdaten", action="store_true",
                    help="leeres Objektblatt statt Testobjekt")
    ap.add_argument("--kostenstellen", metavar="PFAD",
                    help="Kanzlei-Excel mit einem Blatt je Kostenstelle einlesen")
    args = ap.parse_args()

    modell = Modell() if args.ohne_testdaten or args.kostenstellen else testmodell()
    if args.kostenstellen:
        laufende, uebersprungen = lese_kostenstellen(args.kostenstellen, _basisjahr())
        for titel in uebersprungen:
            print(f"übersprungen: Blatt {titel!r} (kein Kostenstellenblatt)")
        modell.objekte = zusammenfuehren(modell.objekte, laufende)
        for lw in laufende:
            print(f"eingelesen: {lw.objekt_id} ({lw.name}) aus Blatt {lw.blatt!r}: "
                  f"Miete {lw.miete:,.2f}, Erhaltung {lw.erhaltung:,.2f}, "
                  f"AfA lt. BWA {lw.abschreibung:,.2f}")
    ziel = Path(args.ausgabe)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    erstelle_mappe(modell).save(ziel)
    print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
