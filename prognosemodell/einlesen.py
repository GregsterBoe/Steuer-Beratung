"""Einleseschicht (Projektplan Abschnitt 16): zwei Eingabedateien werden zum Modell.

Quelle 1, Kostenstellen: die Kanzlei-Excel mit einem Blatt je Kostenstelle, also je
Immobilie, mit Konten und Jahresbeträgen. Sie wird nur gelesen, nie verändert.

Quelle 2, Stammdaten: gepflegt, ein Blatt mit einer Zeile je Objekt (steuerliche
Stammdaten und die Kostenstelle dazu) und ein Blatt Kontenzuordnung, das jedes
Konto einer Kategorie des Modells zuordnet.

Das Layout beider Dateien steht nur hier. Ändert sich das Format der Kanzlei-Excel,
werden die Konstanten unten angepasst, nicht der Rest des Modells.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union

from openpyxl import Workbook, load_workbook

from .modelle import FMT_TEXT, MAX_OBJEKTE, OBJEKT_FELDER, Feld, Modell, Objekt

# --- Quelle 1: Kostenstellenblatt -------------------------------------------------
# Kopfbereich: Beschriftung in Spalte A, Wert in Spalte B; gesucht wird die Beschriftung
KST_TITEL = "Kostenstellenauswertung"
KST_KOPF_KOSTENSTELLE = "Kostenstelle"
KST_KOPF_BEZEICHNUNG = "Bezeichnung"
KST_KOPF_JAHR = "Geschäftsjahr"
KST_KOPF_SUCHZEILEN = 10              # Kopfbereich: so viele Zeilen ab oben
# Kontentabelle: Kopfzeile wird über die Überschrift der Kontospalte gefunden
KST_SPALTE_KONTO = "Konto"
KST_SPALTE_TEXT = "Kontobezeichnung"
KST_SPALTE_BETRAG = "Betrag Geschäftsjahr"

# --- Quelle 2: Stammdaten -----------------------------------------------------------
BLATT_STAMMDATEN = "Objekte"
BLATT_ZUORDNUNG = "Kontenzuordnung"
STAMM_KOPFZEILE = 3                   # Zeile 1 Titel, Zeile 2 Hinweis
ZUORD_KOPFZEILE = 3

# Kategorien der Kontenzuordnung -> Feld im Objektblatt (None = nicht übernehmen)
KATEGORIE_MIETE = "Miete"
KATEGORIE_ERHALTUNG = "Erhaltung"
KATEGORIE_EINNAHMEN = "weitere Einnahmen"
KATEGORIE_AUSGABEN = "weitere Ausgaben"
KATEGORIE_IGNORIEREN = "nicht übernehmen"
KATEGORIEN = {
    KATEGORIE_MIETE: "miete",
    KATEGORIE_ERHALTUNG: "erhaltung",
    KATEGORIE_EINNAHMEN: "weitere_einnahmen",
    KATEGORIE_AUSGABEN: "weitere_ausgaben",
    KATEGORIE_IGNORIEREN: None,
}

# Laufende Werte kommen aus den Kostenstellen, alles andere aus den Stammdaten
LAUFENDE_FELDER = [k for k in KATEGORIEN.values() if k]
STAMM_FELDER = (
    [OBJEKT_FELDER[0],
     Feld("kostenstelle", "Kostenstelle", "", FMT_TEXT, True, 12)]
    + [f for f in OBJEKT_FELDER[1:] if f.key not in LAUFENDE_FELDER]
)
ZUORD_FELDER = [
    Feld("von", "Konto von", "", "0", True, 10),
    Feld("bis", "Konto bis (leer = nur dieses Konto)", "", "0", False, 14),
    Feld("kategorie", "Kategorie", "", FMT_TEXT, True, 18, auswahl=tuple(KATEGORIEN)),
    Feld("bemerkung", "Bemerkung", "", FMT_TEXT, False, 50),
]

Quelle = Union[str, Path, Workbook]


class EinleseFehler(Exception):
    """Eingabedateien unvollständig oder widersprüchlich; meldungen nennt jeden Fund."""

    def __init__(self, meldungen: list):
        self.meldungen = meldungen
        super().__init__("\n".join(meldungen))


@dataclass
class Kostenstelle:
    nummer: str
    bezeichnung: Optional[str]
    jahr: Optional[int]
    blatt: str
    konten: list = field(default_factory=list)   # (Konto, Text, Betrag)


@dataclass
class Ergebnis:
    modell: Modell
    hinweise: list                     # nicht fatal, z. B. Kostenstelle ohne Stammdaten


def _mappe(quelle: Quelle) -> Workbook:
    if isinstance(quelle, Workbook):
        return quelle
    return load_workbook(quelle, data_only=True, read_only=False)


def _text(wert) -> Optional[str]:
    """Zellwert als Text; Kostenstelle 1001 und "1001" sind dasselbe."""
    if wert is None:
        return None
    if isinstance(wert, float) and wert.is_integer():
        wert = int(wert)
    text = str(wert).strip()
    return text or None


def _zahl(wert, wo: str, fehler: list) -> Optional[float]:
    if wert is None or (isinstance(wert, str) and not wert.strip()):
        return None
    if isinstance(wert, bool) or not isinstance(wert, (int, float)):
        fehler.append(f"{wo}: keine Zahl ({wert!r})")
        return None
    return wert


def _kopfzeile(ws, zeile: int, felder, wo: str, fehler: list) -> dict:
    """Spaltennummer je Feld über die Überschrift; Spalten dürfen umgestellt werden."""
    spalten = {}
    for zelle in ws[zeile]:
        text = _text(zelle.value)
        for f in felder:
            if text == f.ueberschrift:
                spalten[f.key] = zelle.column
    for f in felder:
        if f.key not in spalten:
            fehler.append(f"{wo}: Spalte '{f.ueberschrift}' fehlt in Zeile {zeile}")
    return spalten


# --- Quelle 1 ---------------------------------------------------------------------

def _kostenstelle(ws, fehler: list) -> Optional[Kostenstelle]:
    """Ein Kostenstellenblatt lesen; None, wenn das Blatt keins ist (z. B. Anleitung)."""
    kopf = {}
    for zeile in ws.iter_rows(min_row=1, max_row=KST_KOPF_SUCHZEILEN, max_col=2):
        beschriftung = _text(zeile[0].value)
        if beschriftung in (KST_KOPF_KOSTENSTELLE, KST_KOPF_BEZEICHNUNG, KST_KOPF_JAHR):
            kopf[beschriftung] = zeile[1].value
    if KST_KOPF_KOSTENSTELLE not in kopf:
        return None
    wo = f"Kostenstellen, Blatt '{ws.title}'"
    nummer = _text(kopf[KST_KOPF_KOSTENSTELLE])
    if not nummer:
        fehler.append(f"{wo}: Kostenstelle leer")
        return None
    jahr = kopf.get(KST_KOPF_JAHR)
    if not isinstance(jahr, (int, float)) or isinstance(jahr, bool):
        fehler.append(f"{wo}: {KST_KOPF_JAHR} fehlt oder ist keine Zahl")
        jahr = None
    kst = Kostenstelle(nummer, _text(kopf.get(KST_KOPF_BEZEICHNUNG)),
                       int(jahr) if jahr is not None else None, ws.title)

    kopfzeile = next((z[0].row for z in ws.iter_rows(max_col=1)
                      if _text(z[0].value) == KST_SPALTE_KONTO), None)
    if kopfzeile is None:
        fehler.append(f"{wo}: Kontentabelle fehlt (Überschrift '{KST_SPALTE_KONTO}')")
        return kst
    spalten = _kopfzeile(ws, kopfzeile, [Feld("konto", KST_SPALTE_KONTO, "", "", True),
                                         Feld("text", KST_SPALTE_TEXT, "", "", True),
                                         Feld("betrag", KST_SPALTE_BETRAG, "", "", True)],
                         wo, fehler)
    if len(spalten) < 3:
        return kst
    for zeile in range(kopfzeile + 1, ws.max_row + 1):
        konto = ws.cell(row=zeile, column=spalten["konto"]).value
        if konto is None:
            continue                         # Leer- oder Summenzeile ohne Konto
        if not isinstance(konto, (int, float)) or isinstance(konto, bool):
            continue                         # z. B. "Summe"
        betrag = _zahl(ws.cell(row=zeile, column=spalten["betrag"]).value,
                       f"{wo}, Zeile {zeile}", fehler)
        kst.konten.append((int(konto), _text(ws.cell(row=zeile, column=spalten["text"]).value),
                           betrag or 0))
    return kst


def lies_kostenstellen(quelle: Quelle) -> tuple:
    """Alle Kostenstellenblätter: (Liste Kostenstelle, Fehlerliste)."""
    fehler, liste = [], []
    for ws in _mappe(quelle).worksheets:
        kst = _kostenstelle(ws, fehler)
        if kst is not None:
            liste.append(kst)
    gesehen = {}
    for kst in liste:
        if kst.nummer in gesehen:
            fehler.append(f"Kostenstellen: Kostenstelle {kst.nummer} doppelt "
                          f"(Blätter '{gesehen[kst.nummer]}' und '{kst.blatt}')")
        gesehen.setdefault(kst.nummer, kst.blatt)
    if not liste:
        fehler.append("Kostenstellen: kein Blatt mit Kopf "
                      f"'{KST_KOPF_KOSTENSTELLE}' gefunden")
    return liste, fehler


# --- Quelle 2 ---------------------------------------------------------------------

def _zeilen(ws, kopfzeile: int, felder, wo: str, fehler: list):
    """Datenzeilen als (Zeilennummer, {key: Wert}); leere Zeilen werden übersprungen."""
    spalten = _kopfzeile(ws, kopfzeile, felder, wo, fehler)
    if len(spalten) < len(felder):
        return
    for zeile in range(kopfzeile + 1, ws.max_row + 1):
        werte = {f.key: ws.cell(row=zeile, column=spalten[f.key]).value for f in felder}
        if all(w is None or (isinstance(w, str) and not w.strip()) for w in werte.values()):
            continue
        yield zeile, werte


def lies_zuordnung(quelle: Quelle) -> tuple:
    """Kontenzuordnung: (Liste (von, bis, Kategorie), Fehlerliste)."""
    fehler, regeln = [], []
    wb = _mappe(quelle)
    if BLATT_ZUORDNUNG not in wb.sheetnames:
        return [], [f"Stammdaten: Blatt '{BLATT_ZUORDNUNG}' fehlt"]
    for zeile, w in _zeilen(wb[BLATT_ZUORDNUNG], ZUORD_KOPFZEILE, ZUORD_FELDER,
                            f"Stammdaten, Blatt '{BLATT_ZUORDNUNG}'", fehler):
        wo = f"Stammdaten, {BLATT_ZUORDNUNG} Zeile {zeile}"
        von = _zahl(w["von"], wo, fehler)
        bis = _zahl(w["bis"], wo, fehler)
        kategorie = _text(w["kategorie"])
        if von is None:
            fehler.append(f"{wo}: Konto von fehlt")
            continue
        bis = von if bis is None else bis
        if bis < von:
            fehler.append(f"{wo}: Konto bis kleiner als Konto von")
            continue
        if kategorie not in KATEGORIEN:
            fehler.append(f"{wo}: Kategorie {kategorie!r} unbekannt, erlaubt: "
                          + ", ".join(KATEGORIEN))
            continue
        for v, b, _, z in regeln:
            if von <= b and v <= bis:
                fehler.append(f"{wo}: Kontenbereich überschneidet sich mit Zeile {z}")
        regeln.append((int(von), int(bis), kategorie, zeile))
    return [(v, b, k) for v, b, k, _ in regeln], fehler


def lies_stammdaten(quelle: Quelle) -> tuple:
    """Stammdatenzeilen: (Liste {key: Wert}, Fehlerliste)."""
    fehler, saetze = [], []
    wb = _mappe(quelle)
    if BLATT_STAMMDATEN not in wb.sheetnames:
        return [], [f"Stammdaten: Blatt '{BLATT_STAMMDATEN}' fehlt"]
    for zeile, w in _zeilen(wb[BLATT_STAMMDATEN], STAMM_KOPFZEILE, STAMM_FELDER,
                            f"Stammdaten, Blatt '{BLATT_STAMMDATEN}'", fehler):
        wo = f"Stammdaten, {BLATT_STAMMDATEN} Zeile {zeile}"
        satz = {"_zeile": zeile}
        for f in STAMM_FELDER:
            if f.format == FMT_TEXT:
                satz[f.key] = _text(w[f.key])
            else:
                satz[f.key] = _zahl(w[f.key], f"{wo}, {f.ueberschrift}", fehler)
        if satz["kaufjahr"] is not None:
            satz["kaufjahr"] = int(satz["kaufjahr"])
        for f in STAMM_FELDER:
            if f.pflicht and satz[f.key] is None:
                fehler.append(f"{wo}: {f.ueberschrift} fehlt")
        saetze.append(satz)
    for key, text in (("objekt_id", "ObjektID"), ("kostenstelle", "Kostenstelle")):
        zeilen = {}
        for s in saetze:
            if s[key] is not None:
                zeilen.setdefault(s[key], []).append(s["_zeile"])
        for wert, z in zeilen.items():
            if len(z) > 1:
                fehler.append(f"Stammdaten: {text} {wert} doppelt (Zeilen "
                              + ", ".join(map(str, z)) + ")")
    return saetze, fehler


# --- Zusammenführen ---------------------------------------------------------------

def _kategorie(konto: int, regeln) -> Optional[str]:
    return next((k for v, b, k in regeln if v <= konto <= b), None)


def lies_modell(stammdaten: Quelle, kostenstellen: Quelle) -> Ergebnis:
    """Beide Eingabedateien zu einem Modell; EinleseFehler mit allen Funden sonst.

    Basisjahr ist das Geschäftsjahr der Kostenstellen; es muss überall gleich sein.
    Verkäufe und Neuobjekte bleiben leer, sie werden in der Mappe geplant.
    """
    kst_liste, fehler = lies_kostenstellen(kostenstellen)
    regeln, f = lies_zuordnung(stammdaten)
    fehler += f
    saetze, f = lies_stammdaten(stammdaten)
    fehler += f
    hinweise = []

    jahre = sorted({k.jahr for k in kst_liste if k.jahr is not None})
    if len(jahre) > 1:
        fehler.append("Kostenstellen: unterschiedliche Geschäftsjahre "
                      + ", ".join(map(str, jahre)) + "; alle Blätter müssen dasselbe Jahr zeigen")

    kst_nach_nummer = {k.nummer: k for k in kst_liste}
    benutzt = set()
    objekte = []
    for s in saetze:
        wo = f"Stammdaten, {BLATT_STAMMDATEN} Zeile {s['_zeile']} ({s['objekt_id']})"
        kst = kst_nach_nummer.get(s["kostenstelle"])
        if kst is None:
            if s["kostenstelle"] is not None:
                fehler.append(f"{wo}: Kostenstelle {s['kostenstelle']} nicht in der "
                              "Kostenstellendatei")
            continue
        benutzt.add(kst.nummer)
        summen = dict.fromkeys(LAUFENDE_FELDER, 0)
        for konto, text, betrag in kst.konten:
            kategorie = _kategorie(konto, regeln)
            if kategorie is None:
                if betrag:
                    fehler.append(f"Kostenstellen, Blatt '{kst.blatt}': Konto {konto} "
                                  f"({text}) ohne Kontenzuordnung")
                continue
            if KATEGORIEN[kategorie]:
                summen[KATEGORIEN[kategorie]] += betrag
        werte = {f.key: s[f.key] for f in STAMM_FELDER if f.key != "kostenstelle"}
        werte["name"] = werte["name"] or kst.bezeichnung
        for key in ("weitere_einnahmen", "weitere_ausgaben"):
            summen[key] = summen[key] or None        # 0 bleibt im Objektblatt leer
        objekte.append(Objekt(**werte, **summen))

    for kst in kst_liste:
        if kst.nummer not in benutzt:
            hinweise.append(f"Kostenstelle {kst.nummer} ({kst.bezeichnung}, Blatt "
                            f"'{kst.blatt}') hat keine Stammdaten und wird nicht übernommen")
    if len(objekte) > MAX_OBJEKTE:
        fehler.append(f"Stammdaten: {len(objekte)} Objekte, die Mappe fasst {MAX_OBJEKTE}")
    if not saetze:
        fehler.append(f"Stammdaten: keine Objekte im Blatt '{BLATT_STAMMDATEN}'")
    if fehler:
        raise EinleseFehler(fehler)
    parameter = {"par_Basisjahr": jahre[0]} if jahre else {}
    return Ergebnis(Modell(objekte=objekte, parameter=parameter), hinweise)
