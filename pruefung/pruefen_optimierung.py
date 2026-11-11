"""Optimierung der Browser-Fassung (browser/optimierung.js) gegen die Sollwerte der
Makrofälle „Optimierung: …“ aus pruefung.pruefen prüfen: dieselben Modelle, dieselben
Einstellungen, dieselben erwarteten Hebel.

Aufruf: python -m pruefung.pruefen_optimierung   (braucht Node.js und LibreOffice)

Die Eingabe entsteht wie im Abgleich aus dem Modell des Falls und der gerechneten Mappe.
"""

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from prognosemodell.browser import metadaten
from prognosemodell.modelle import NEU_FELDER

from . import abgleich, pruefen

CLI = abgleich.WURZEL / "browser" / "optimiere_cli.js"

# Fall -> (Einstellungen abweichend von der Vorgabe, Prüfungen auf das Ergebnis)
FAELLE = {
    "Optimierung: günstiger Kredit, Kaufjahr, Varianten": ({}, [
        ("Kreditquote NEU-001", lambda e: e["neuobjekte"][0]["felder"]["kredit_quote"], 0.8),
        ("Kaufjahr NEU-001", lambda e: e["neuobjekte"][0]["felder"]["kaufjahr"], 2029),
        ("§ 6b nur bis QNG-Grenze bleibt leer", lambda e: e["neuobjekte"][0]["felder"]["qng_grenze"], ""),
        ("Quote vorher", lambda e: e["lauf"]["zeilen"][0]["vorQuote"], 0.4),
        ("Kaufjahr vorher", lambda e: e["lauf"]["zeilen"][0]["vorJahr"], 2030),
        ("§ 6b-Schalter vorher", lambda e: e["lauf"]["zeilen"][0]["vor6b"], "nein"),
        ("Endvermögen vorher = Variante Ausgang", lambda e: e["lauf"]["vorher"], lambda e: e["ausgang"]["A"]),
        ("Endvermögen nachher = Variante Ergebnis", lambda e: e["lauf"]["nachher"], lambda e: e["variante"]["A"]),
        ("Endvermögen nachher höher", lambda e: e["lauf"]["nachher"] > e["lauf"]["vorher"], True),
        ("Nebenbedingungen erfüllt", lambda e: e["lauf"]["erfuellt"], True),
        ("geänderte Hebel", lambda e: e["lauf"]["aenderungen"], 2),
        ("Kreditquoten Ausgang", lambda e: e["ausgang"]["kreditquoten"], "NEU-001 40 %"),
        ("Kreditquoten Ergebnis", lambda e: e["variante"]["kreditquoten"], "NEU-001 80 %"),
        ("Prüfung", lambda e: e["variante"]["pruefung"], "OK"),
        ("Eingabe des Aufrufers unverändert", lambda e: e["unveraendert"], True),
    ]),
    "Optimierung: Ausschüttung bestimmt die Kreditquote": (
        {"hebelJahr": "nein", "kreditSchritt": 0.1, "festhalten": "nein"}, [
            ("Kreditquote NEU-001", lambda e: e["neuobjekte"][0]["felder"]["kredit_quote"], 0.6),
            ("Kredit NEU-001", lambda e: e["neuobjekte"][0]["berechnet"]["kredit"], 360_000),
            ("Kaufjahr NEU-001", lambda e: e["neuobjekte"][0]["felder"]["kaufjahr"], 2035),
            ("Lauf erfüllt", lambda e: e["lauf"]["erfuellt"], True),
            ("Nebenbedingungen erfüllt", lambda e: e["kennzahlen"]["erfuellt"], True),
            ("tragbare Ausschüttung", lambda e: e["tragbar"], 19_000),
            ("Parameter Ausschüttung bleibt", lambda e: e["ausschuettung"], 20_000),
            ("geänderte Hebel", lambda e: e["lauf"]["aenderungen"], 1),
        ]),
    "Optimierung: § 6b nur bis QNG-Grenze mit zweitem Neuobjekt": (
        {"hebelJahr": "nein", "hebelKredit": "nein", "festhalten": "nein"}, [
            ("§ 6b-Schalter NEU-001", lambda e: e["neuobjekte"][0]["felder"]["qng_grenze"], "ja"),
            ("§ 6b-Schalter NEU-002 bleibt leer", lambda e: e["neuobjekte"][1]["felder"]["qng_grenze"], ""),
            ("ü1 NEU-001", lambda e: e["neuobjekte"][0]["berechnet"]["ue1"], 80_000),
            ("§ 6b-Schalter NEU-002 vorher", lambda e: e["lauf"]["zeilen"][1]["vor6b"], "nein"),
            ("geänderte Hebel", lambda e: e["lauf"]["aenderungen"], 1),
        ]),
}


def meta() -> dict:
    pruefungen = metadaten()["pruefungen"]
    return {"fehlerKeys": [p["key"] for p in pruefungen if p["art"] == "Fehler"],
            "neuFelder": [f.key for f in NEU_FELDER], "pruefungen": pruefungen}


def gleich(ist, soll) -> bool:
    if isinstance(soll, bool) or isinstance(ist, bool) or isinstance(soll, str):
        return ist == soll
    return isinstance(ist, (int, float)) and abs(ist - soll) < 0.005


def main() -> int:
    node = shutil.which("node")
    if not node:
        sys.exit("Node.js (node) nicht gefunden.")
    namen = [f[0] for f in pruefen._faelle()]
    auftraege = []
    with tempfile.TemporaryDirectory() as tmp:
        for fall, (einstellungen, _) in FAELLE.items():
            i = namen.index(fall)
            modell, zellen = abgleich.modell_zu(i)
            wb = pruefen.durchrechnen(modell, Path(tmp) / f"fall{i}", zellen,
                                      profil=Path(tmp) / "lo-profil")
            ein = abgleich.eingabe(modell, wb)
            ein.pop("_leer_kst", None)
            auftraege.append({"eingabe": ein, "einstellungen": einstellungen, "meta": meta()})
    lauf = subprocess.run([node, str(CLI)], input=json.dumps(auftraege), capture_output=True,
                          text=True, check=True, timeout=300)
    fehler = 0
    for (fall, (_, pruefungen)), ergebnis in zip(FAELLE.items(), json.loads(lauf.stdout)):
        print(fall)
        if "fehler" in ergebnis:
            print("  Abbruch:", ergebnis["fehler"])
            fehler += 1
            continue
        for text, ist, soll in pruefungen:
            wert = ist(ergebnis)
            ziel = soll(ergebnis) if callable(soll) else soll
            ok = gleich(wert, ziel)
            fehler += not ok
            print(f"  {'OK    ' if ok else 'FEHLER'} {text}: {wert!r}" + ("" if ok else f", erwartet {ziel!r}"))
        print(f"  Rechenläufe: {ergebnis['lauf']['laeufe']}, {ergebnis['lauf']['meldung']}")
    print("\n" + ("alle Prüfungen bestanden" if not fehler else f"{fehler} Abweichung(en)"))
    return 1 if fehler else 0


if __name__ == "__main__":
    raise SystemExit(main())
