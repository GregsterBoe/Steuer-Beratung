"""Datenmodelle: Parameter, Objektfelder und Objekt-Stammdaten.

Hier steht nur, *was* in die Mappe gehört. Wie es geschrieben wird, regelt
mappe.py, die Formeln liefert formeln.py.
"""

from dataclasses import dataclass, field
from typing import Optional, Union

# Zahlenformate
FMT_EURO = '#,##0 "€"'
FMT_PROZENT = "0.00%"
FMT_JAHR = "0"
FMT_ZAHL = "0"
FMT_TEXT = "@"

# Feste Zeilenzahl des Objektblatts; benannte Bereiche laufen über alle Zeilen.
MAX_OBJEKTE = 200


@dataclass(frozen=True)
class Parameter:
    name: str              # benannter Bereich, z. B. par_Basisjahr
    bezeichnung: str
    wert: Union[float, int, str]  # Zahl, Text oder Formel ("=...")
    format: str
    erlaeuterung: str
    auswahl: Optional[tuple] = None  # Dropdown-Werte


# Reihenfolge = Reihenfolge auf dem Parameterblatt.
# Alle Sätze sind Platzhalter und vor dem Echteinsatz fachlich zu prüfen.
PARAMETER = [
    Parameter("par_Basisjahr", "Basisjahr (Ist)", 2026, FMT_JAHR,
              "Letztes Ist-Jahr; Stammdaten beziehen sich auf dessen Ende"),
    Parameter("par_Prognosejahre", "Prognosejahre", 20, FMT_ZAHL,
              "Länge des Jahresrasters"),
    Parameter("par_Startjahr", "erstes Prognosejahr", "=par_Basisjahr+1", FMT_JAHR,
              "berechnet"),
    Parameter("par_Endjahr", "letztes Prognosejahr", "=par_Basisjahr+par_Prognosejahre",
              FMT_JAHR, "berechnet"),
    Parameter("par_Steuerwelt", "Steuerwelt", "GmbH", FMT_TEXT,
              "MVP rechnet nur die gewerbliche Struktur mit GmbH", ("GmbH",)),
    Parameter("par_Steuersatz", "Grenzsteuersatz", 0.30, FMT_PROZENT,
              "KSt, SolZ und GewSt zusammen; Platzhalter, prüfen"),
    Parameter("par_Mietsteig", "Mietsteigerung p. a.", 0.02, FMT_PROZENT, "Platzhalter"),
    Parameter("par_Erhaltsteig", "Erhaltungssteigerung p. a.", 0.025, FMT_PROZENT,
              "Platzhalter"),
    Parameter("par_Kostensteig", "Steigerung weitere Ausgaben p. a.", 0.02, FMT_PROZENT,
              "Platzhalter; weitere Einnahmen wachsen mit der Mietsteigerung"),
    Parameter("par_Wertsteig", "Wertsteigerung p. a.", 0.02, FMT_PROZENT, "Platzhalter"),
    Parameter("par_GrESt", "Grunderwerbsteuersatz", 0.05, FMT_PROZENT,
              "abhängig vom Bundesland; Platzhalter"),
    Parameter("par_Alternativrendite", "Rendite Alternativanlage p. a.", 0.04, FMT_PROZENT,
              "für Szenario B; Platzhalter"),
    Parameter("par_Aufteilung", "Erlösaufteilung Standard", "Buchwert", FMT_TEXT,
              "Aufteilung Verkaufserlös auf Gebäude und G+B", ("Buchwert", "Verkehrswert")),
    Parameter("par_6bVorbesitz", "§ 6b Mindest-Vorbesitzzeit (Jahre)", 6, FMT_ZAHL, ""),
    Parameter("par_6bFrist", "§ 6b Reinvestitionsfrist (Jahre)", 4, FMT_ZAHL, ""),
    Parameter("par_6bZuschlag", "§ 6b Gewinnzuschlag je Jahr", 0.06, FMT_PROZENT,
              "bei Auflösung ohne Reinvestition"),
    Parameter("par_Szenario", "aktives Szenario", "A", FMT_TEXT,
              "A = § 6b-Kette, B = sofort versteuern", ("A", "B")),
]


@dataclass(frozen=True)
class Feld:
    """Eine Spalte des Objektblatts."""
    key: str
    ueberschrift: str
    name: str            # benannter Bereich über die Spalte
    format: str
    pflicht: bool
    breite: int = 16
    minimum: Optional[float] = None  # Datenvalidierung
    maximum: Optional[float] = None
    ganzzahl: bool = False


OBJEKT_FELDER = [
    Feld("objekt_id", "ObjektID", "obj_ID", FMT_TEXT, True, 12),
    Feld("name", "Objektname", "obj_Name", FMT_TEXT, False, 28),
    Feld("ak_gebaeude", "AK Gebäude", "obj_AKGebaeude", FMT_EURO, True, minimum=0),
    Feld("ak_gub", "AK G+B", "obj_AKGuB", FMT_EURO, True, minimum=0),
    Feld("kaufjahr", "Kaufjahr", "obj_Kaufjahr", FMT_JAHR, True, 10,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("afa_satz", "AfA-Satz", "obj_AfASatz", FMT_PROZENT, True, 10, minimum=0, maximum=0.2),
    Feld("restbuchwert", "Restbuchwert Gebäude Basisjahr", "obj_Restbuchwert", FMT_EURO, True,
         minimum=0),
    Feld("verkehrswert", "Verkehrswert aktuell", "obj_Verkehrswert", FMT_EURO, False, minimum=0),
    Feld("vk_quote_gebaeude", "Verkehrswertanteil Gebäude", "obj_VKQuoteGeb", FMT_PROZENT, False,
         minimum=0, maximum=1),
    Feld("miete", "Miete Basisjahr", "obj_MieteBasis", FMT_EURO, True, minimum=0),
    Feld("erhaltung", "Erhaltung Basisjahr", "obj_ErhBasis", FMT_EURO, True, minimum=0),
    Feld("weitere_einnahmen", "weitere Einnahmen Basisjahr", "obj_EinnBasis", FMT_EURO, False,
         minimum=0),
    Feld("weitere_ausgaben", "weitere Ausgaben Basisjahr", "obj_AusgBasis", FMT_EURO, False,
         minimum=0),
]

# Berechnete Statusspalte direkt nach den Eingabefeldern
STATUS_UEBERSCHRIFT = "Status"
STATUS_NAME = "obj_Status"


@dataclass(frozen=True)
class Spalte:
    """Eine Spalte des Prognoseblatts."""
    key: str
    ueberschrift: str
    name: str            # benannter Bereich über die Spalte
    format: str
    breite: int = 14


# Blatt Prognose: eine Zeile je Objekt und Jahr (Long-Format, Projektplan Abschnitt 8)
PROGNOSE_SPALTEN = [
    Spalte("id", "ObjektID", "prg_ID", FMT_TEXT, 12),
    Spalte("jahr", "Jahr", "prg_Jahr", FMT_JAHR, 8),
    Spalte("aktiv", "aktiv", "prg_Aktiv", FMT_ZAHL, 8),
    Spalte("miete", "Miete", "prg_Miete", FMT_EURO),
    Spalte("einnahmen", "weitere Einnahmen", "prg_Einnahmen", FMT_EURO),
    Spalte("erhaltung", "Erhaltung", "prg_Erhaltung", FMT_EURO),
    Spalte("ausgaben", "weitere Ausgaben", "prg_Ausgaben", FMT_EURO),
    Spalte("afa", "AfA Gebäude", "prg_AfA", FMT_EURO),
    Spalte("buchwert", "Buchwert Gebäude Ende", "prg_Buchwert", FMT_EURO, 16),
    Spalte("ergebnis", "Ergebnis vor Finanzierung", "prg_Ergebnis", FMT_EURO, 16),
]


def prognosejahre() -> int:
    """Länge des Jahresrasters; legt die Zeilen je Objekt im Prognoseblatt fest."""
    return next(p.wert for p in PARAMETER if p.name == "par_Prognosejahre")


@dataclass
class Objekt:
    """Stammdaten eines Bestandsobjekts. None = Feld leer lassen."""
    objekt_id: Optional[str]
    name: Optional[str] = None
    ak_gebaeude: Optional[float] = None
    ak_gub: Optional[float] = None
    kaufjahr: Optional[int] = None
    afa_satz: Optional[float] = None
    restbuchwert: Optional[float] = None
    verkehrswert: Optional[float] = None
    vk_quote_gebaeude: Optional[float] = None
    miete: Optional[float] = None
    erhaltung: Optional[float] = None
    weitere_einnahmen: Optional[float] = None
    weitere_ausgaben: Optional[float] = None


@dataclass
class Modell:
    objekte: list = field(default_factory=list)
