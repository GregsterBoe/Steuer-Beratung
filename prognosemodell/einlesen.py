"""Einleseschicht für die Kanzlei-Excel mit einem Blatt je Kostenstelle.

Jedes Blatt folgt der DATEV-BWA Form 01: Spalte B trägt die BWA-Zeilennummer,
die Kopfzeile die Spaltenjahre. Gelesen wird die Jahresspalte des Basisjahrs,
gesucht wird über die BWA-Nummer, nie über die Zeilenposition.

Ändert sich das Quellformat, wird nur dieses Modul angepasst (Projektplan, Abschnitt 16).
"""

import dataclasses
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from .modelle import Objekt

# Feste Stellen im Blatt
ZELLE_ID = "B2"        # Kostenstelle, z. B. "KSt 1"
ZELLE_NAME = "C2"      # Objektbezeichnung, z. B. "KC 24+26"
KOPFZEILE = 4          # B "Nr.", danach Jahres- und Monatsspalten
SPALTE_NR = 2          # B

# BWA-Zeilen je Modellfeld; fehlende Zeilen zählen als 0
BWA_MIETE = (1020,)
BWA_EINNAHMEN = (1090,)
BWA_ERHALTUNG = (1250,)
BWA_AUSGABEN = (1100, 1120, 1140, 1150, 1160, 1180, 1200, 1220, 1260)
BWA_ABSCHREIBUNG = (1240,)  # nur Abgleich mit der AfA-Fortschreibung


class EinleseFehler(ValueError):
    pass


@dataclass
class LaufendeWerte:
    """Werte eines Kostenstellenblatts im Basisjahr."""
    blatt: str
    objekt_id: str
    name: Optional[str]
    miete: float
    weitere_einnahmen: float
    erhaltung: float
    weitere_ausgaben: float
    abschreibung: float


def _jahresspalte(ws, basisjahr: int) -> int:
    for zelle in ws[KOPFZEILE]:
        if isinstance(zelle.value, (int, float)) and not isinstance(zelle.value, bool) \
                and int(zelle.value) == basisjahr:
            return zelle.column
    raise EinleseFehler(f"Blatt {ws.title!r}: keine Spalte {basisjahr} in Zeile {KOPFZEILE}")


def _bwa_zeilen(ws) -> dict:
    zeilen = {}
    for zeile in range(KOPFZEILE + 1, ws.max_row + 1):
        nr = ws.cell(row=zeile, column=SPALTE_NR).value
        if isinstance(nr, (int, float)):
            zeilen[int(nr)] = zeile
    return zeilen


def _summe(ws, ws_formeln, zeilen: dict, nummern: tuple, spalte: int) -> float:
    summe = 0.0
    for nr in nummern:
        if nr not in zeilen:
            continue
        wert = ws.cell(row=zeilen[nr], column=spalte).value
        if wert is None:
            formel = ws_formeln.cell(row=zeilen[nr], column=spalte).value
            if isinstance(formel, str) and formel.startswith("="):
                raise EinleseFehler(
                    f"Blatt {ws.title!r}, BWA {nr}: Formel ohne berechneten Wert. "
                    "Datei in Excel öffnen, speichern und erneut einlesen.")
            continue
        if not isinstance(wert, (int, float)):
            raise EinleseFehler(f"Blatt {ws.title!r}, BWA {nr}: kein Zahlenwert ({wert!r})")
        summe += wert
    return summe


def lese_kostenstellen(pfad, basisjahr: int) -> list:
    """Liest alle Blätter der Kanzlei-Excel; je Blatt eine LaufendeWerte-Zeile."""
    pfad = Path(pfad)
    werte_wb = load_workbook(pfad, data_only=True)   # berechnete Werte
    formel_wb = load_workbook(pfad, data_only=False)  # nur für Fehlermeldungen
    ergebnis, gesehen = [], {}
    for ws in werte_wb.worksheets:
        objekt_id = ws[ZELLE_ID].value
        if objekt_id is None or str(objekt_id).strip() == "":
            raise EinleseFehler(f"Blatt {ws.title!r}: keine Kostenstelle in {ZELLE_ID}")
        objekt_id = str(objekt_id).strip()
        if objekt_id in gesehen:
            raise EinleseFehler(
                f"Kostenstelle {objekt_id!r} doppelt (Blätter {gesehen[objekt_id]!r} "
                f"und {ws.title!r})")
        gesehen[objekt_id] = ws.title
        name = ws[ZELLE_NAME].value
        spalte = _jahresspalte(ws, basisjahr)
        zeilen = _bwa_zeilen(ws)
        wf = formel_wb[ws.title]

        def summe(nummern):
            return _summe(ws, wf, zeilen, nummern, spalte)

        ergebnis.append(LaufendeWerte(
            blatt=ws.title,
            objekt_id=objekt_id,
            name=str(name).strip() if name is not None else None,
            miete=summe(BWA_MIETE),
            weitere_einnahmen=summe(BWA_EINNAHMEN),
            erhaltung=summe(BWA_ERHALTUNG),
            weitere_ausgaben=summe(BWA_AUSGABEN),
            abschreibung=summe(BWA_ABSCHREIBUNG),
        ))
    return ergebnis


def zusammenfuehren(stammdaten: list, laufende: list) -> list:
    """Laufende Werte in die Objekte übernehmen (Schlüssel: ObjektID).

    Kostenstellen ohne Stammdaten werden als neue Objekte angelegt; deren
    steuerliche Pflichtfelder bleiben leer und die Statusspalte meldet sie.
    Objekte ohne Kostenstellenblatt bleiben unverändert.
    """
    nach_id = {o.objekt_id: o for o in stammdaten}
    ergebnis = list(stammdaten)
    for lw in laufende:
        felder = dict(miete=lw.miete, weitere_einnahmen=lw.weitere_einnahmen,
                      erhaltung=lw.erhaltung, weitere_ausgaben=lw.weitere_ausgaben)
        if lw.objekt_id in nach_id:
            alt = nach_id[lw.objekt_id]
            neu = dataclasses.replace(alt, name=alt.name or lw.name, **felder)
            ergebnis[ergebnis.index(alt)] = neu
        else:
            ergebnis.append(Objekt(objekt_id=lw.objekt_id, name=lw.name, **felder))
    return ergebnis
