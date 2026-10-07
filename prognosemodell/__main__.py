"""Mappe generieren: python -m prognosemodell [--ausgabe PFAD] [--ohne-testdaten]
[--kostenstellen PFAD] [--inventar PFAD [--inventar-stand JAHR]] [--makros] [--schnellcheck]"""

import argparse
import sys
from pathlib import Path

from .einlesen import (anlagen_zusammenfuehren, lese_inventar, lese_kostenstellen,
                       ordne_anlagen_zu, stand_aus_dateiname, zusammenfuehren)
from .mappe import erstelle_mappe
from .modelle import PARAMETER, ZUORDNUNG_KOST1, Modell
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
    ap.add_argument("--inventar", metavar="PFAD",
                    help="Anlagenverzeichnis (DATEV Inventarübersicht) einlesen; KOST1 = Kostenstelle")
    ap.add_argument("--inventar-stand", metavar="JAHR", type=int,
                    help="Wirtschaftsjahr der Buchwerte im Anlagenverzeichnis; Standard: Jahr im "
                         "Dateinamen, sonst Basisjahr − 1")
    ap.add_argument("--makros", action="store_true",
                    help="als .xlsm mit VBA-Steuerung speichern (braucht LibreOffice beim Bauen)")
    ap.add_argument("--schnellcheck", action="store_true",
                    help="schlanke Mappe: wenige Eingaben, zentrale Annahmen, Empfehlung")
    args = ap.parse_args()

    modell = Modell() if args.ohne_testdaten or args.kostenstellen else testmodell()
    stand = args.inventar_stand or (stand_aus_dateiname(args.inventar) if args.inventar else None)
    if stand is not None and stand != _basisjahr() - 1:
        modell.parameter["par_AnlStand"] = stand
    laufende = []
    if args.kostenstellen:
        laufende, uebersprungen = lese_kostenstellen(args.kostenstellen, _basisjahr(), stand)
        for titel, grund in uebersprungen:
            print(f"übersprungen: Blatt {titel!r} ({grund})")
        neukauf = [lw for lw in laufende if lw.neukauf]
        laufende = [lw for lw in laufende if not lw.neukauf]
        modell.objekte = zusammenfuehren(modell.objekte, laufende)
        modell.kostenstellen = {lw.objekt_id: lw for lw in laufende}
        modell.neukauf = {lw.objekt_id: lw for lw in neukauf}
        for lw in neukauf:
            jahre = sorted({j for werte in lw.jahre.values() for j in werte})
            print(f"Neukauf-Kostenstelle: {lw.objekt_id} ({lw.name}) aus Blatt {lw.blatt!r}"
                  + (f", Planwerte {jahre[0]}–{jahre[-1]}" if jahre else ", ohne Planwerte"))
        for lw in laufende:
            print(f"eingelesen: {lw.objekt_id} ({lw.name}) aus Blatt {lw.blatt!r}: "
                  f"Miete {lw.miete:,.2f}, Erhaltung {lw.erhaltung:,.2f}, "
                  f"AfA lt. BWA {lw.abschreibung:,.2f}"
                  + (f", {len(lw.anlagen)} Anlagengruppe(n)" if lw.anlagen else "")
                  + (f", AfA-Plan {min(lw.afa_plan)}–{max(lw.afa_plan)} "
                     f"({len(lw.afa_plan)} Jahre)" if lw.afa_plan else ""))
    inventar = []
    if args.inventar:
        inventar, abgang = lese_inventar(args.inventar)
        inventar = ordne_anlagen_zu(inventar, modell.objekte)
        ids = {o.objekt_id for o in modell.objekte}
        mit_objekt = sum(a.objekt_id in ids for a in inventar)
        print(f"Anlagenverzeichnis: {len(inventar)} Anlagen eingelesen ({abgang} abgegangen), "
              f"{mit_objekt} einem Objekt zugeordnet, Stand {stand or _basisjahr() - 1}")
        for a in inventar:
            if a.zuordnung and a.zuordnung != ZUORDNUNG_KOST1:
                ziel = f"-> {a.objekt_id}" if a.objekt_id else "-> kein Objekt"
                print(f"  ohne KOST1: {a.nr} „{a.bezeichnung}“ {ziel} ({a.zuordnung}), "
                      f"im Blatt Anlagen prüfen")
    modell.anlagen, abgleich = anlagen_zusammenfuehren(inventar, laufende)
    for objekt_id, bw_inventar, bw_kst, afa_kst in abgleich:
        hinweis = "" if abs(bw_inventar - bw_kst) < 1 else "  <- weicht ab, Zuordnung prüfen"
        print(f"Abgleich {objekt_id}: Buchwert abnutzbar lt. Anlagenverzeichnis "
              f"{bw_inventar:,.2f}, lt. Kostenstellenblatt {bw_kst:,.2f}{hinweis}")
    modell.schnellcheck = args.schnellcheck
    standard = Path(ap.get_default("ausgabe"))
    if args.schnellcheck:
        standard = standard.with_name("Schnellcheck_VV.xlsx")
        if args.ausgabe == ap.get_default("ausgabe"):
            args.ausgabe = str(standard)
    ziel = Path(args.ausgabe)
    if ziel.is_dir():  # nur Ordner angegeben: Standarddateiname darin
        ziel = ziel / standard.name
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
