"""Prüfskript für die Einleseschicht: synthetische Kostenstellenblätter im Musterlayout.

Aufruf: python -m pruefung.pruefen_einlesen
Läuft ohne LibreOffice. Beendet sich mit Fehlercode 1 bei Abweichung.
"""

import dataclasses
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

from prognosemodell.einlesen import EinleseFehler, lese_kostenstellen, zusammenfuehren
from prognosemodell.testdaten import testobjekt
from prognosemodell.vorlagen import erstelle_vorlage

TOLERANZ = 0.01

# BWA-Nr. -> Wert in der Jahresspalte 2026 (frei erfunden, nur Layout wie im Muster)
WERTE = {
    1020: 120_000.0, 1051: 120_000.0, 1090: 1_500.0,
    1100: 2_000.0, 1120: 10_000.0, 1140: 3_000.0, 1150: 4_000.0,
    1240: 16_000.0, 1250: -500.0, 1260: 1_000.0,
    1280: 36_000.0, 1310: 9_000.0,
}


def kostenstellenblatt(ws, kst: str, objekt: str, werte: dict, formel_statt_wert=None):
    """Blatt im Layout des Musters: B2/C2 Kopf, Zeile 4 Spaltenköpfe, ab Zeile 6 BWA-Zeilen."""
    ws["B2"], ws["C2"] = kst, objekt
    ws["B4"], ws["C4"] = "Nr.", "Bezeichnung kurz"
    ws["F4"], ws["S4"] = 2025, 2026  # G–R wären die Monate, T ff. Folgejahre
    for zeile, nr in enumerate(sorted(werte), start=6):
        ws.cell(row=zeile, column=2, value=nr)
        ws.cell(row=zeile, column=19,
                value="=SUM(G{0}:R{0})".format(zeile) if nr == formel_statt_wert
                else werte[nr])


def schreibe(pfad: Path, blaetter: list, annahmen: bool = False) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    if annahmen:  # Fremdblätter: B2 belegt; einmal ohne "Nr.", einmal mit, aber ohne Jahr
        ws = wb.create_sheet("Annahmen")
        ws["B2"], ws["B4"], ws["C4"] = "Mietsteigerung", "Jahr", 2025
        ws = wb.create_sheet("Übersicht")
        ws["B2"], ws["B4"], ws["C4"] = "Summe", "Nr.", "Bezeichnung"
    for titel, kst, objekt, werte, formel in blaetter:
        kostenstellenblatt(wb.create_sheet(titel), kst, objekt, werte, formel)
    wb.save(pfad)
    return pfad


def main() -> int:
    fehler = 0

    def pruefe(fall, ist, soll):
        nonlocal fehler
        gut = abs(ist - soll) <= TOLERANZ if isinstance(soll, float) else ist == soll
        print(f"{'OK    ' if gut else 'FEHLER'}  {fall}: Soll {soll!r}, Ist {ist!r}")
        fehler += not gut

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        ohne_1090 = {k: v for k, v in WERTE.items() if k != 1090}
        pfad = schreibe(tmp / "kst.xlsx", [
            ("KSt 1", "KSt 1", "KC 24+26", WERTE, None),
            ("KSt 2", "KSt 2", "Objekt B", ohne_1090, None),
        ], annahmen=True)
        (lw1, lw2), uebersprungen = lese_kostenstellen(pfad, 2026)
        pruefe("Blätter mit anderem Format werden mit Grund übersprungen", uebersprungen, [
            ("Annahmen", "kein 'Nr.' in Zeile 4, Spalte B"),
            ("Übersicht", "keine Spalte 2026 in Zeile 4")])
        pruefe("ID aus B2", lw1.objekt_id, "KSt 1")
        pruefe("Name aus C2", lw1.name, "KC 24+26")
        pruefe("Miete = BWA 1020", lw1.miete, 120_000.0)
        pruefe("weitere Einnahmen = BWA 1090", lw1.weitere_einnahmen, 1_500.0)
        pruefe("Erhaltung = BWA 1250 (auch negativ)", lw1.erhaltung, -500.0)
        pruefe("weitere Ausgaben = 1100–1220 + 1260, ohne AfA/Erhaltung/Zins",
               lw1.weitere_ausgaben, 20_000.0)
        pruefe("AfA lt. BWA 1240", lw1.abschreibung, 16_000.0)
        pruefe("fehlende BWA-Zeile zählt 0", lw2.weitere_einnahmen, 0.0)

        try:
            lese_kostenstellen(pfad, 2030)
            pruefe("Basisjahr in keinem Blatt meldet Fehler", "kein Fehler", "EinleseFehler")
        except EinleseFehler:
            pruefe("Basisjahr in keinem Blatt meldet Fehler", "EinleseFehler", "EinleseFehler")

        nur_annahmen = schreibe(tmp / "annahmen.xlsx", [], annahmen=True)
        try:
            lese_kostenstellen(nur_annahmen, 2026)
            pruefe("ohne Kostenstellenblatt meldet Fehler", "kein Fehler", "EinleseFehler")
        except EinleseFehler:
            pruefe("ohne Kostenstellenblatt meldet Fehler", "EinleseFehler", "EinleseFehler")

        doppelt = schreibe(tmp / "doppelt.xlsx", [
            ("A", "KSt 1", "x", WERTE, None), ("B", "KSt 1", "y", WERTE, None)])
        try:
            lese_kostenstellen(doppelt, 2026)
            pruefe("doppelte Kostenstelle meldet Fehler", "kein Fehler", "EinleseFehler")
        except EinleseFehler:
            pruefe("doppelte Kostenstelle meldet Fehler", "EinleseFehler", "EinleseFehler")

        ungerechnet = schreibe(tmp / "formel.xlsx", [("A", "KSt 1", "x", WERTE, 1020)])
        try:
            lese_kostenstellen(ungerechnet, 2026)
            pruefe("Formel ohne Wert meldet Fehler", "kein Fehler", "EinleseFehler")
        except EinleseFehler:
            pruefe("Formel ohne Wert meldet Fehler", "EinleseFehler", "EinleseFehler")

        # Zielformat der Planungsreferenz: Kopf "Jahr 2026", Monatsspalten als Datum,
        # Planspalten "Plan 2027", 2028 …; Summenblatt "Alle Objekte" wird übersprungen
        vorlage = tmp / "vorlage.xlsx"
        erstelle_vorlage(basisjahr=2026).save(vorlage)
        (v1, v2), uebersprungen = lese_kostenstellen(vorlage, 2026)
        pruefe("Vorlage: Summenblatt übersprungen", uebersprungen,
               [("Alle Objekte", "Summenblatt aller Kostenstellen")])
        pruefe("Vorlage: Kostenstellen", (v1.objekt_id, v2.objekt_id), ("KSt 1", "KSt 2"))
        pruefe("Vorlage: Name aus C2", v1.name, "Musterstraße 1")
        pruefe("Vorlage: Miete aus Spalte 'Jahr 2026', nicht aus Vorjahr oder Monat",
               v1.miete, 60_000.0)
        pruefe("Vorlage: Erhaltung", v1.erhaltung, 8_000.0)
        pruefe("Vorlage: weitere Ausgaben 1140 + 1150 + 1260", v1.weitere_ausgaben, 3_600.0)
        pruefe("Vorlage: weitere Einnahmen KSt 2", v2.weitere_einnahmen, 1_500.0)
        pruefe("Vorlage: AfA lt. BWA KSt 2", v2.abschreibung, 30_000.0)
        try:
            (p1, _), _ = lese_kostenstellen(vorlage, 2027)
            pruefe("Vorlage: Planspalte 'Plan 2027' ist leer, Werte 0", p1.miete, 0.0)
        except EinleseFehler as e:
            pruefe("Vorlage: Planspalte 'Plan 2027' lesbar", str(e), "kein Fehler")
        try:
            lese_kostenstellen(vorlage, 2031)
            pruefe("Vorlage: Planjahr 2031 als Zahl erkannt", "kein Fehler", "kein Fehler")
        except EinleseFehler as e:
            pruefe("Vorlage: Planjahr 2031 als Zahl erkannt", str(e), "kein Fehler")

        stamm = testobjekt()  # OBJ-001
        stamm_kst1 = dataclasses.replace(stamm, objekt_id="KSt 1", name=None)
        objekte = zusammenfuehren([stamm_kst1, stamm], [lw1, lw2])
        pruefe("Zusammenführen: Anzahl Objekte", len(objekte), 3)
        pruefe("Zusammenführen: Miete überschrieben", objekte[0].miete, 120_000.0)
        pruefe("Zusammenführen: Stammdaten bleiben", objekte[0].ak_gebaeude, 800_000)
        pruefe("Zusammenführen: Objekt ohne Blatt unverändert", objekte[1].miete, 60_000)
        pruefe("Zusammenführen: neue Kostenstelle ohne AK", objekte[2].ak_gebaeude, None)

    print(f"\n{fehler} Abweichung(en)" if fehler else "\nAlle Prüfungen bestanden.")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main())
