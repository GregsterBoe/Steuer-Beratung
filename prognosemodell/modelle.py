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
    Parameter("par_AfADegressiv", "AfA degressiv (Neuobjekte)", 0.05, FMT_PROZENT,
              "§ 7 Abs. 5a EStG, Wohngebäude mit Baubeginn 10/2023 bis 9/2029; vom Restbuchwert"),
    Parameter("par_Alternativrendite", "Rendite Alternativanlage p. a.", 0.04, FMT_PROZENT,
              "für Szenario B; Platzhalter"),
    Parameter("par_6bVorbesitz", "§ 6b Mindest-Vorbesitzzeit (Jahre)", 6, FMT_ZAHL, ""),
    Parameter("par_6bFrist", "§ 6b Reinvestitionsfrist (Jahre)", 4, FMT_ZAHL, ""),
    Parameter("par_6bFristNeubau", "§ 6b Frist bei Neubau (Jahre)", 6, FMT_ZAHL,
              "wenn mit dem Neubau vor Ende der Regelfrist begonnen wurde (§ 6b Abs. 3)"),
    Parameter("par_6bZuschlag", "§ 6b Gewinnzuschlag je Jahr", 0.06, FMT_PROZENT,
              "bei Auflösung ohne Reinvestition"),
    Parameter("par_Szenario", "aktives Szenario", "A", FMT_TEXT,
              "A = § 6b-Kette, B = sofort versteuern", ("A", "B")),
]


@dataclass(frozen=True)
class Feld:
    """Eine Spalte eines Eingabeblatts (Objekte, Verkäufe)."""
    key: str
    ueberschrift: str
    name: str            # benannter Bereich über die Spalte
    format: str
    pflicht: bool
    breite: int = 16
    minimum: Optional[float] = None  # Datenvalidierung
    maximum: Optional[float] = None
    ganzzahl: bool = False
    auswahl: Optional[tuple] = None  # Dropdown-Werte


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
    """Eine berechnete Spalte (Prognose, berechneter Teil der Verkäufe)."""
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
    # Wertentwicklung bei Halten, unabhängig vom Verkauf (Baseline der Übersicht)
    Spalte("verkehrswert", "Verkehrswert Ende", "prg_Verkehrswert", FMT_EURO, 16),
    # 1 = Objekt am Jahresende noch im Bestand; im Verkaufsjahr schon 0 (Verkauf zum Jahresende)
    Spalte("bestand", "im Bestand Ende", "prg_Bestand", FMT_ZAHL, 10),
    # 1 = Neuobjekt (Etappe 6); zählt nur im Plan, nicht in der Baseline
    Spalte("neu", "Neuobjekt", "prg_Neu", FMT_ZAHL, 10),
    # Etappe 7: Buchwerte für stille Reserven und Baseline
    Spalte("buchwert_gub", "Buchwert G+B Ende", "prg_BuchwertGuB", FMT_EURO),
    # Gebäudebuchwert, als würde das Objekt nie verkauft (nur Bestandsobjekte, sonst 0)
    Spalte("buchwert_halten", "Buchwert Gebäude bei Halten", "prg_BuchwertHalten", FMT_EURO, 16),
]

# Blatt Verkäufe (Etappe 4, Projektplan Abschnitt 11): Eingaben, dann berechnete Spalten
MAX_VERKAEUFE = 50
VERKAUF_FELDER = [
    Feld("objekt_id", "ObjektID", "vk_ID", FMT_TEXT, True, 14),
    Feld("jahr", "Verkaufsjahr", "vk_Jahr", FMT_JAHR, True, 12,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("preis", "Verkaufspreis", "vk_Preis", FMT_EURO, True, minimum=0),
    Feld("kosten", "Verkaufskosten", "vk_Kosten", FMT_EURO, False, minimum=0),
    Feld("anteil_gub", "Anteil G+B lt. Kaufvertrag", "vk_AnteilGuBVertrag", FMT_PROZENT, False,
         minimum=0, maximum=1),
    Feld("nutzung_6b", "§ 6b nutzen", "vk_6b", FMT_TEXT, False, 10, auswahl=("ja", "nein")),
    # verlängert die Reinvestitionsfrist auf par_6bFristNeubau
    Feld("neubau_6b", "§ 6b Neubau begonnen", "vk_6bNeubau", FMT_TEXT, False, 11,
         auswahl=("ja", "nein")),
]
VERKAUF_SPALTEN = [
    Spalte("vorbesitz", "Vorbesitzzeit Jahre", "vk_Vorbesitz", FMT_ZAHL, 11),
    Spalte("buchwert_geb", "Buchwert Gebäude Ende Verkaufsjahr", "vk_BuchwertGeb", FMT_EURO, 16),
    Spalte("ak_gub", "AK G+B", "vk_AKGuB", FMT_EURO),
    Spalte("nettoerloes", "Nettoerlös", "vk_Nettoerloes", FMT_EURO),
    Spalte("quote_gub", "Anteil G+B verwendet", "vk_AnteilGuB", FMT_PROZENT, 11),
    Spalte("erloes_geb", "Erlösanteil Gebäude", "vk_ErloesGeb", FMT_EURO),
    Spalte("erloes_gub", "Erlösanteil G+B", "vk_ErloesGuB", FMT_EURO),
    Spalte("gewinn_geb", "Gewinn Gebäude", "vk_GewinnGeb", FMT_EURO),
    Spalte("gewinn_gub", "Gewinn G+B", "vk_GewinnGuB", FMT_EURO),
    Spalte("gewinn", "Veräußerungsgewinn", "vk_Gewinn", FMT_EURO, 16),
]
VERKAUF_STATUS_NAME = "vk_Status"
# Statustexte, auf die der Rücklagenspiegel zugreift
STATUS_OK = "OK"
# Gewinn ist gültig berechnet, wird aber sofort versteuert
STATUS_6B_UNZULAESSIG = "§ 6b unzulässig: Vorbesitzzeit zu kurz"

# Blatt Rücklagen (Etappe 5, Projektplan Abschnitt 12)
# Teil 1: je Verkaufszeile eine Rücklagenzeile, gefüllt nur bei § 6b ja und Status OK
RUECKLAGE_SPALTEN = [
    Spalte("id", "RücklageID", "rl_ID", FMT_TEXT, 14),
    Spalte("objekt_id", "ObjektID Herkunft", "rl_ObjektID", FMT_TEXT, 12),
    Spalte("jahr", "Bildungsjahr (Verkaufsjahr)", "rl_Jahr", FMT_JAHR, 11),
    Spalte("geb", "Rücklage Gebäude", "rl_Geb", FMT_EURO),
    Spalte("gub", "Rücklage G+B", "rl_GuB", FMT_EURO),
    Spalte("betrag", "Rücklage gesamt", "rl_Betrag", FMT_EURO),
    Spalte("fristjahr", "Fristjahr", "rl_Fristjahr", FMT_JAHR, 9),
    Spalte("ueb_geb", "übertragen Gebäude", "rl_UebGeb", FMT_EURO),
    Spalte("ueb_gub", "übertragen G+B", "rl_UebGuB", FMT_EURO),
    Spalte("aufloesung", "Auflösung im Fristjahr", "rl_Aufloesung", FMT_EURO),
    Spalte("zuschlag", "Gewinnzuschlag", "rl_Zuschlag", FMT_EURO),
    Spalte("hinweis", "Hinweis", "rl_Hinweis", FMT_TEXT, 30),
]
# Teil 2: Spiegel je Prognosejahr
RUECKLAGE_JAHR_SPALTEN = [
    Spalte("jahr", "Jahr", "rls_Jahr", FMT_JAHR, 8),
    Spalte("gewinne", "Veräußerungsgewinne", "rls_Gewinne", FMT_EURO, 15),
    Spalte("bildung", "Einstellung in Rücklage", "rls_Bildung", FMT_EURO, 15),
    Spalte("uebertragung", "Übertragung auf Neuobjekte", "rls_Uebertragung", FMT_EURO, 15),
    Spalte("aufloesung", "Auflösung", "rls_Aufloesung", FMT_EURO),
    Spalte("zuschlag", "Gewinnzuschlag", "rls_Zuschlag", FMT_EURO),
    Spalte("steuerwirksam", "steuerwirksam aus Verkauf und Rücklage", "rls_Steuerwirksam",
           FMT_EURO, 17),
    Spalte("bestand", "Rücklagenbestand Ende", "rls_Bestand", FMT_EURO, 15),
]


AFA_LINEAR = "linear"
AFA_DEGRESSIV = "degressiv"

# Blatt Neuobjekte (Etappe 6, Projektplan Abschnitt 13): Kauf zum Jahresende,
# Miete und AfA ab dem Folgejahr
MAX_NEUOBJEKTE = 50
NEU_FELDER = [
    Feld("neu_id", "NeuID", "ne_ID", FMT_TEXT, True, 12),
    Feld("name", "Name", "ne_Name", FMT_TEXT, False, 22),
    Feld("kaufjahr", "Kaufjahr", "ne_Kaufjahr", FMT_JAHR, True, 10,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("kaufpreis", "Kaufpreis", "ne_Kaufpreis", FMT_EURO, True, minimum=0),
    Feld("anteil_gub", "Anteil G+B", "ne_AnteilGuB", FMT_PROZENT, True, 10, minimum=0, maximum=1),
    Feld("nebenkosten", "Kaufnebenkosten", "ne_Nebenkosten", FMT_EURO, False, minimum=0),
    Feld("afa_satz", "AfA-Satz", "ne_AfASatz", FMT_PROZENT, True, 10, minimum=0, maximum=0.2),
    # leer = linear; degressiv: par_AfADegressiv vom Restbuchwert, Wechsel zur linearen AfA,
    # sobald Restbuchwert / Restnutzungsdauer höher ist; Nutzungsdauer = 1 / AfA-Satz
    Feld("afa_methode", "AfA-Methode", "ne_AfAMethode", FMT_TEXT, False, 11,
         auswahl=(AFA_LINEAR, AFA_DEGRESSIV)),
    Feld("mietrendite", "Mietrendite auf Kaufpreis", "ne_Mietrendite", FMT_PROZENT, False, 11,
         minimum=0, maximum=1),
    Feld("erhaltungsquote", "Erhaltung auf Kaufpreis", "ne_ErhQuote", FMT_PROZENT, False, 11,
         minimum=0, maximum=1),
    Feld("quelle", "Quelle RücklageID", "ne_Quelle", FMT_TEXT, False, 16),
]
NEU_SPALTEN = [
    # 1 = rechnet in der Prognose mit (Pflichtfelder da, ID eindeutig, Kaufjahr im Raster)
    Spalte("gueltig", "im Modell", "ne_Gueltig", FMT_ZAHL, 8),
    Spalte("ak_gub_neu", "AK G+B neu", "ne_AKGuBNeu", FMT_EURO),
    Spalte("ak_geb_neu", "AK Gebäude neu", "ne_AKGebNeu", FMT_EURO),
    Spalte("rl_geb", "Rücklage Gebäude verfügbar", "ne_RLGeb", FMT_EURO),
    Spalte("rl_gub", "Rücklage G+B verfügbar", "ne_RLGuB", FMT_EURO),
    Spalte("ue1", "ü1 Gebäudegewinn auf Gebäude", "ne_Ue1", FMT_EURO),
    Spalte("ue2", "ü2 G+B-Gewinn auf G+B", "ne_Ue2", FMT_EURO),
    Spalte("ue3", "ü3 G+B-Gewinn auf Gebäude", "ne_Ue3", FMT_EURO),
    Spalte("ue_gesamt", "übertragen gesamt", "ne_UeGesamt", FMT_EURO),
    Spalte("afa_basis", "AfA-Basis Gebäude", "ne_AfABasis", FMT_EURO, 16),
    Spalte("ak_gub", "steuerliche AK G+B", "ne_AKGuB", FMT_EURO),
]
NEU_STATUS_NAME = "ne_Status"


# Blatt Liquidität (Etappe 7, Projektplan Abschnitt 14): je Jahr eine Zeile,
# links der Plan, rechts die Baseline „alles halten“ (eine Spalte Abstand)
LIQ_SPALTEN = [
    Spalte("jahr", "Jahr", "liq_Jahr", FMT_JAHR, 8),
    Spalte("ergebnis", "laufendes Ergebnis", "liq_Ergebnis", FMT_EURO),
    Spalte("verkauf", "steuerwirksam aus Verkauf und Rücklage", "liq_Verkauf", FMT_EURO, 16),
    Spalte("zve", "Ergebnis vor Verlustvortrag", "liq_ZvE", FMT_EURO),
    Spalte("vortrag_genutzt", "Verlustvortrag genutzt", "liq_VortragGenutzt", FMT_EURO),
    Spalte("bemessung", "Bemessungsgrundlage", "liq_Bemessung", FMT_EURO),
    Spalte("vortrag", "Verlustvortrag Ende", "liq_Vortrag", FMT_EURO),
    Spalte("steuer", "Steuer", "liq_Steuer", FMT_EURO),
    Spalte("einnahmen", "Mieten und weitere Einnahmen", "liq_Einnahmen", FMT_EURO),
    Spalte("ausgaben", "Erhaltung und weitere Ausgaben", "liq_Ausgaben", FMT_EURO),
    Spalte("verkaufserloes", "Verkaufserlöse netto", "liq_Verkaufserloes", FMT_EURO),
    Spalte("rueckfluss", "davon Buchwert-Rückfluss", "liq_Rueckfluss", FMT_EURO),
    Spalte("kauf", "Kauf Neuobjekte inkl. Nebenkosten", "liq_Kauf", FMT_EURO, 15),
    Spalte("zufluss", "freier Mittelzufluss", "liq_Zufluss", FMT_EURO),
    Spalte("kum", "Liquidität kumuliert Ende", "liq_Kum", FMT_EURO, 15),
]
LIQ_BASIS_SPALTEN = [
    Spalte("jahr", "Jahr", "lqb_Jahr", FMT_JAHR, 8),
    Spalte("einnahmen", "Mieten und weitere Einnahmen", "lqb_Einnahmen", FMT_EURO),
    Spalte("ausgaben", "Erhaltung und weitere Ausgaben", "lqb_Ausgaben", FMT_EURO),
    Spalte("afa", "AfA Gebäude", "lqb_AfA", FMT_EURO),
    Spalte("ergebnis", "laufendes Ergebnis", "lqb_Ergebnis", FMT_EURO),
    Spalte("vortrag_genutzt", "Verlustvortrag genutzt", "lqb_VortragGenutzt", FMT_EURO),
    Spalte("bemessung", "Bemessungsgrundlage", "lqb_Bemessung", FMT_EURO),
    Spalte("vortrag", "Verlustvortrag Ende", "lqb_Vortrag", FMT_EURO),
    Spalte("steuer", "Steuer", "lqb_Steuer", FMT_EURO),
    Spalte("zufluss", "freier Mittelzufluss", "lqb_Zufluss", FMT_EURO),
    Spalte("kum", "Liquidität kumuliert Ende", "lqb_Kum", FMT_EURO, 15),
]

# Blatt Auswertung (Etappe 7): Kennzahlen je Jahr, links Plan, rechts Baseline
AUSWERTUNG_SPALTEN = [
    Spalte("jahr", "Jahr", "aus_Jahr", FMT_JAHR, 8),
    Spalte("ergebnis", "laufendes Ergebnis", "aus_Ergebnis", FMT_EURO),
    Spalte("verkauf", "steuerwirksam aus Verkauf und Rücklage", "aus_Verkauf", FMT_EURO, 16),
    Spalte("guv", "Gesamt-GuV vor Steuern", "aus_GuV", FMT_EURO),
    Spalte("steuer", "Steuer", "aus_Steuer", FMT_EURO),
    Spalte("nach_steuer", "Ergebnis nach Steuern", "aus_NachSteuer", FMT_EURO),
    Spalte("steuer_kum", "Steuer kumuliert", "aus_SteuerKum", FMT_EURO),
    Spalte("verkehrswert", "Verkehrswert Bestand", "aus_Verkehrswert", FMT_EURO, 15),
    Spalte("buchwert", "Buchwert Bestand (Gebäude + G+B)", "aus_Buchwert", FMT_EURO, 16),
    Spalte("stille_reserven", "stille Reserven", "aus_StilleReserven", FMT_EURO),
    Spalte("ruecklage", "§ 6b-Rücklage Bestand", "aus_Ruecklage", FMT_EURO),
    Spalte("vortrag", "Verlustvortrag", "aus_Vortrag", FMT_EURO),
    Spalte("liquiditaet", "Liquidität kumuliert", "aus_Liquiditaet", FMT_EURO),
    Spalte("vermoegen", "Gesamtvermögen vor latenter Steuer", "aus_Vermoegen", FMT_EURO, 16),
    Spalte("latente_steuer", "latente Steuer", "aus_LatenteSteuer", FMT_EURO),
    Spalte("vermoegen_netto", "Gesamtvermögen nach latenter Steuer", "aus_VermoegenNetto",
           FMT_EURO, 17),
]
AUSWERTUNG_BASIS_SPALTEN = [
    Spalte("jahr", "Jahr", "asb_Jahr", FMT_JAHR, 8),
    Spalte("verkehrswert", "Verkehrswert Bestand", "asb_Verkehrswert", FMT_EURO, 15),
    Spalte("buchwert", "Buchwert Bestand (Gebäude + G+B)", "asb_Buchwert", FMT_EURO, 16),
    Spalte("stille_reserven", "stille Reserven", "asb_StilleReserven", FMT_EURO),
    Spalte("vortrag", "Verlustvortrag", "asb_Vortrag", FMT_EURO),
    Spalte("liquiditaet", "Liquidität kumuliert", "asb_Liquiditaet", FMT_EURO),
    Spalte("vermoegen", "Gesamtvermögen vor latenter Steuer", "asb_Vermoegen", FMT_EURO, 16),
    Spalte("latente_steuer", "latente Steuer", "asb_LatenteSteuer", FMT_EURO),
    Spalte("vermoegen_netto", "Gesamtvermögen nach latenter Steuer", "asb_VermoegenNetto",
           FMT_EURO, 17),
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
class Verkauf:
    """Geplanter Verkauf zum Ende des Verkaufsjahrs. None = Feld leer lassen."""
    objekt_id: str
    jahr: Optional[int]
    preis: Optional[float] = None
    kosten: Optional[float] = None
    anteil_gub: Optional[float] = None   # Kaufvertrag; leer = Verkehrswertanteil aus Objekte
    nutzung_6b: Optional[str] = None
    neubau_6b: Optional[str] = None      # ja = Frist par_6bFristNeubau statt par_6bFrist


@dataclass
class Neuobjekt:
    """Reinvestitionsobjekt, gekauft zum Ende des Kaufjahrs. None = Feld leer lassen."""
    neu_id: str
    kaufjahr: Optional[int]
    kaufpreis: Optional[float] = None
    anteil_gub: Optional[float] = None
    afa_satz: Optional[float] = None
    name: Optional[str] = None
    afa_methode: Optional[str] = None    # "linear" (leer) oder "degressiv"
    nebenkosten: Optional[float] = None
    mietrendite: Optional[float] = None
    erhaltungsquote: Optional[float] = None
    quelle: Optional[str] = None         # RücklageID, z. B. "RL-OBJ-001"


@dataclass
class Modell:
    objekte: list = field(default_factory=list)
    verkaeufe: list = field(default_factory=list)
    neuobjekte: list = field(default_factory=list)
