"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten]
[--kostenstellen PFAD] [--makros]"""

import argparse
import sys
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
    ap.add_argument("--makros", action="store_true",
                    help="als .xlsm mit VBA-Steuerung speichern (braucht LibreOffice beim Bauen)")
    args = ap.parse_args()

    modell = Modell() if args.ohne_testdaten or args.kostenstellen else testmodell()
    if args.kostenstellen:
        laufende, uebersprungen = lese_kostenstellen(args.kostenstellen, _basisjahr())
        for titel, grund in uebersprungen:
            print(f"übersprungen: Blatt {titel!r} ({grund})")
        modell.objekte = zusammenfuehren(modell.objekte, laufende)
        modell.kostenstellen = {lw.objekt_id: lw for lw in laufende}
        for lw in laufende:
            print(f"eingelesen: {lw.objekt_id} ({lw.name}) aus Blatt {lw.blatt!r}: "
                  f"Miete {lw.miete:,.2f}, Erhaltung {lw.erhaltung:,.2f}, "
                  f"AfA lt. BWA {lw.abschreibung:,.2f}")
    ziel = Path(args.ausgabe)
    if ziel.is_dir():  # nur Ordner angegeben: Standarddateiname darin
        ziel = ziel / Path(ap.get_default("ausgabe")).name
    if args.makros or ziel.suffix.lower() == ".xlsm":
        ziel = ziel.with_suffix(".xlsm")
    ziel.parent.mkdir(parents=True, exist_ok=True)
    try:
        if ziel.suffix == ".xlsm":
            from .makros import speichere_mit_makros
            speichere_mit_makros(erstelle_mappe(modell), ziel)
        else:
            erstelle_mappe(modell).save(ziel)
    except PermissionError:
        sys.exit(f"Kann {ziel} nicht schreiben. Ist die Datei noch in Excel geöffnet? "
                 "Bitte schließen oder mit --ausgabe einen anderen Dateinamen angeben.")
    print(f"geschrieben: {ziel}")


if __name__ == "__main__":
    main()
