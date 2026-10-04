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
MAX_VERKAEUFE = MAX_OBJEKTE  # höchstens ein Verkauf je Objekt
MAX_NEUOBJEKTE = 50


@dataclass(frozen=True)
class Parameter:
    name: str              # benannter Bereich, z. B. par_Basisjahr
    bezeichnung: str
    wert: Union[float, int, str]  # Zahl, Text oder Formel ("=...")
    format: str
    erlaeuterung: str
    auswahl: Optional[tuple] = None  # Dropdown-Werte
    minimum: Optional[float] = None  # Datenvalidierung (Dezimalzahl)
    maximum: Optional[float] = None


STEUERWELT_GMBH = "GmbH"
SZENARIO_A, SZENARIO_B, SZENARIO_C = "A", "B", "C"
SZENARIEN = (SZENARIO_A, SZENARIO_B, SZENARIO_C)
AUFTEILUNGEN = ("Buchwert", "Verkehrswert")
STEUERWELTEN = (STEUERWELT_GMBH, "Privat / GbR vermögensverwaltend",
                "gewerblich (Personengesellschaft)")

# Reihenfolge = Reihenfolge auf dem Parameterblatt.
# Alle Sätze sind Platzhalter und vor dem Echteinsatz fachlich zu prüfen.
PARAMETER = [
    Parameter("par_Basisjahr", "Basisjahr (Ist)", 2026, FMT_JAHR,
              "Letztes Ist-Jahr; Stammdaten beziehen sich auf dessen Ende"),
    Parameter("par_Prognosejahre", "Prognosejahre", 20, FMT_ZAHL,
              "Länge des Jahresrasters; legt der Generator fest, eine Änderung "
              "in Excel verlängert die Prognose nicht"),
    Parameter("par_Startjahr", "erstes Prognosejahr", "=par_Basisjahr+1", FMT_JAHR,
              "berechnet"),
    Parameter("par_Endjahr", "letztes Prognosejahr", "=par_Basisjahr+par_Prognosejahre",
              FMT_JAHR, "berechnet"),
    Parameter("par_Steuerwelt", "Steuerwelt", STEUERWELT_GMBH, FMT_TEXT,
              "Rechtsform; MVP rechnet nur GmbH (siehe Projektplan Abschnitt 18)",
              STEUERWELTEN),
    Parameter("par_StatusSteuerwelt", "Status Steuerwelt",
              f'=IF(par_Steuerwelt="{STEUERWELT_GMBH}","OK",'
              f'"nicht im MVP – Ergebnisse gelten nur für GmbH")',
              FMT_TEXT, "berechnet"),
    Parameter("par_Steuersatz", "Grenzsteuersatz", 0.30, FMT_PROZENT,
              "KSt, SolZ und GewSt zusammen; Platzhalter, prüfen"),
    Parameter("par_Mietsteig", "Mietsteigerung p. a.", 0.02, FMT_PROZENT,
              "ab Basisjahr, erstes Prognosejahr schon gesteigert; negativ = Rückgang; "
              "Platzhalter", minimum=-0.1, maximum=0.2),
    Parameter("par_Erhaltsteig", "Erhaltungssteigerung p. a.", 0.025, FMT_PROZENT,
              "wie Mietsteigerung; Platzhalter", minimum=-0.1, maximum=0.2),
    Parameter("par_Wertsteig", "Wertsteigerung p. a.", 0.02, FMT_PROZENT, "Platzhalter",
              minimum=-0.1, maximum=0.2),
    Parameter("par_GrESt", "Grunderwerbsteuersatz", 0.05, FMT_PROZENT,
              "abhängig vom Bundesland; Platzhalter"),
    Parameter("par_Alternativrendite", "Rendite Alternativanlage p. a.", 0.04, FMT_PROZENT,
              "vor Steuern, auf die freien Mittel in beiden Szenarien; der Zinsertrag wird mit "
              "dem Grenzsteuersatz besteuert; Platzhalter", minimum=-0.1, maximum=0.2),
    Parameter("par_Aufteilung", "Erlösaufteilung Standard", "Buchwert", FMT_TEXT,
              "Aufteilung Verkaufserlös auf Gebäude und G+B; gilt, wenn im Blatt Verkäufe "
              "keine Methode steht", AUFTEILUNGEN),
    Parameter("par_6bVorbesitz", "§ 6b Mindest-Vorbesitzzeit (Jahre)", 6, FMT_ZAHL,
              "gezählt als Verkaufsjahr minus Kaufjahr (Verkauf zum Jahresende)"),
    Parameter("par_6bFrist", "§ 6b Reinvestitionsfrist (Jahre)", 4, FMT_ZAHL,
              "Fristjahr = Verkaufsjahr + Frist; Neubau-Verlängerung auf 6 Jahre hier eintragen"),
    Parameter("par_6bZuschlag", "§ 6b Gewinnzuschlag je Jahr", 0.06, FMT_PROZENT,
              "je volles Jahr des Bestehens, bei Auflösung ohne Reinvestition"),
    Parameter("par_Szenario", "aktives Szenario", SZENARIO_A, FMT_TEXT,
              "A = § 6b-Kette wie erfasst; B = jeder Veräußerungsgewinn sofort versteuert, "
              "Neuobjekte mit Quelle-Rücklage entfallen, das Kapital bleibt in der "
              "Alternativanlage; C = sofort versteuert, Neuobjekte trotzdem gekauft, ohne "
              "Übertrag mit voller AfA-Basis", SZENARIEN),
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
    auswahl: Optional[tuple] = None  # Dropdown-Werte
    auswahl_bereich: Optional[str] = None  # Dropdown aus einem benannten Bereich


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
class Rechenspalte:
    """Eine Formelspalte (Prognose, berechneter Teil der Verkäufe)."""
    key: str
    ueberschrift: str
    name: str            # benannter Bereich über die Spalte
    format: str
    breite: int = 14
    summe: bool = False  # Jahresblätter: in der Summenzeile aufaddieren


PrognoseSpalte = Rechenspalte

# Blatt Prognose, Long-Format: Zeile je Objekt und Jahr
PROGNOSE_SPALTEN = [
    PrognoseSpalte("id", "ObjektID", "prg_ID", FMT_TEXT, 12),
    PrognoseSpalte("jahr", "Jahr", "prg_Jahr", FMT_JAHR, 8),
    PrognoseSpalte("aktiv", "aktiv", "prg_Aktiv", FMT_ZAHL, 7),
    PrognoseSpalte("miete", "Miete", "prg_Miete", FMT_EURO),
    PrognoseSpalte("erhaltung", "Erhaltung", "prg_Erhaltung", FMT_EURO),
    PrognoseSpalte("afa", "AfA", "prg_AfA", FMT_EURO),
    PrognoseSpalte("buchwert", "Buchwert Gebäude Ende", "prg_Buchwert", FMT_EURO, 16),
    PrognoseSpalte("ergebnis", "Ergebnis vor Finanzierung", "prg_Ergebnis", FMT_EURO, 16),
    PrognoseSpalte("bestand", "im Bestand Jahresende", "prg_Bestand", FMT_ZAHL, 9),
    PrognoseSpalte("bw_gub", "Buchwert G+B Ende", "prg_BuchwertGuB", FMT_EURO),
    PrognoseSpalte("bw_gesamt", "Buchwert gesamt Ende", "prg_BuchwertGesamt", FMT_EURO, 16),
    PrognoseSpalte("verkehrswert", "Verkehrswert Ende", "prg_Verkehrswert", FMT_EURO, 16),
    PrognoseSpalte("stille_reserven", "stille Reserven", "prg_StilleReserven", FMT_EURO, 16),
]


# Blatt Verkäufe: erst die Eingaben, dann Status und Formeln
VERKAUF_FELDER = [
    Feld("objekt_id", "ObjektID", "vk_ID", FMT_TEXT, True, 12),
    Feld("jahr", "Verkaufsjahr", "vk_Jahr", FMT_JAHR, True, 10,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("preis", "Verkaufspreis", "vk_Preis", FMT_EURO, False, minimum=0),
    Feld("faktor", "oder Faktor × Jahresmiete", "vk_Faktor", "0.0", False, 12,
         minimum=0, maximum=100),
    Feld("kosten", "Verkaufskosten", "vk_Kosten", FMT_EURO, False, minimum=0),
    Feld("aufteilung", "Aufteilung (leer = Parameter)", "vk_Aufteilung", FMT_TEXT, False, 14,
         auswahl=AUFTEILUNGEN),
    Feld("nutzung_6b", "6b-Nutzung", "vk_6b", FMT_TEXT, True, 10, auswahl=("ja", "nein")),
]
VERKAUF_STATUS_NAME = "vk_Status"

VERKAUF_SPALTEN = [
    Rechenspalte("preis_angesetzt", "Verkaufspreis angesetzt", "vk_PreisAngesetzt", FMT_EURO, 16),
    Rechenspalte("bw_geb", "Buchwert Gebäude Verkaufsjahr", "vk_BuchwertGeb", FMT_EURO, 16),
    Rechenspalte("bw_gesamt", "Buchwert gesamt (mit G+B)", "vk_BuchwertGesamt", FMT_EURO, 16),
    Rechenspalte("erloes_geb", "Erlösanteil Gebäude", "vk_ErloesGeb", FMT_EURO, 16),
    Rechenspalte("gewinn_geb", "Gewinn Gebäude", "vk_GewinnGeb", FMT_EURO, 16),
    Rechenspalte("gewinn_gub", "Gewinn G+B", "vk_GewinnGuB", FMT_EURO, 16),
    Rechenspalte("gewinn", "Veräußerungsgewinn gesamt", "vk_Gewinn", FMT_EURO, 16),
]


# Blatt Rücklagen, Teil 1: je Zeile im Blatt Verkäufe eine Zeile (gleiche Zeilennummer)
RUECKLAGE_6B_GEBILDET = "Rücklage gebildet"
RUECKLAGE_SPALTEN = [
    Rechenspalte("id", "ObjektID", "rl_ID", FMT_TEXT, 12),
    Rechenspalte("jahr", "Verkaufsjahr", "rl_Jahr", FMT_JAHR, 10),
    Rechenspalte("kaufjahr", "Kaufjahr", "rl_Kaufjahr", FMT_JAHR, 10),
    Rechenspalte("vorbesitz", "Vorbesitzzeit Jahre", "rl_Vorbesitz", FMT_ZAHL, 10),
    Rechenspalte("status", "Status § 6b", "rl_Status", FMT_TEXT, 22),
    Rechenspalte("ruecklage_id", "RücklageID", "rl_RuecklageID", FMT_TEXT, 14),
    Rechenspalte("gewinn", "Veräußerungsgewinn", "rl_Gewinn", FMT_EURO, 16),
    Rechenspalte("betrag_geb", "Rücklage Gebäude", "rl_BetragGeb", FMT_EURO, 16),
    Rechenspalte("betrag_gub", "Rücklage G+B", "rl_BetragGuB", FMT_EURO, 16),
    Rechenspalte("ruecklage", "Rücklage gesamt", "rl_Ruecklage", FMT_EURO, 16),
    Rechenspalte("steuerpflichtig", "steuerpflichtig im Verkaufsjahr", "rl_Steuerpflichtig",
                 FMT_EURO, 16),
    Rechenspalte("fristjahr", "Fristjahr", "rl_Fristjahr", FMT_JAHR, 10),
    Rechenspalte("uebertrag_geb", "übertragen aus Gebäude-Rücklage", "rl_UebertragGeb",
                 FMT_EURO, 16),
    Rechenspalte("uebertrag_gub", "übertragen aus G+B-Rücklage", "rl_UebertragGuB",
                 FMT_EURO, 16),
    Rechenspalte("rest", "Restrücklage = Auflösung im Fristjahr", "rl_Rest", FMT_EURO, 18),
    Rechenspalte("zuschlag", "Gewinnzuschlag", "rl_Zuschlag", FMT_EURO, 16),
]

# Blatt Rücklagen, Teil 2: Spiegel je Prognosejahr
RUECKLAGE_JAHR_SPALTEN = [
    Rechenspalte("jahr", "Jahr", "rlj_Jahr", FMT_JAHR, 8),
    Rechenspalte("gewinn", "Veräußerungsgewinne", "rlj_Gewinn", FMT_EURO, 16),
    Rechenspalte("bildung", "Bildung Rücklage", "rlj_Bildung", FMT_EURO, 16),
    Rechenspalte("uebertrag", "Übertrag auf Neuobjekte", "rlj_Uebertrag", FMT_EURO, 14),
    Rechenspalte("aufloesung", "Auflösung Fristablauf", "rlj_Aufloesung", FMT_EURO, 16),
    Rechenspalte("zuschlag", "Gewinnzuschlag", "rlj_Zuschlag", FMT_EURO, 14),
    Rechenspalte("stand", "Rücklage Stand Jahresende", "rlj_Stand", FMT_EURO, 16),
    Rechenspalte("steuerpflichtig", "steuerpflichtig aus Verkauf und Auflösung",
                 "rlj_Steuerpflichtig", FMT_EURO, 18),
    Rechenspalte("steuer", "Steuer darauf", "rlj_Steuer", FMT_EURO, 14),
]


# Blatt Neuobjekte: Reinvestitionsobjekte, Eingaben, Status, dann Formeln
NEU_FELDER = [
    Feld("neu_id", "NeuID", "neu_ID", FMT_TEXT, True, 12),
    Feld("name", "Objektname", "neu_Name", FMT_TEXT, False, 24),
    Feld("kaufjahr", "Kaufjahr", "neu_Kaufjahr", FMT_JAHR, True, 10,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("kaufpreis", "Kaufpreis", "neu_Kaufpreis", FMT_EURO, True, minimum=0),
    Feld("anteil_gub", "Anteil G+B", "neu_AnteilGuB", FMT_PROZENT, True, 10,
         minimum=0, maximum=1),
    Feld("nebenkosten", "Kaufnebenkosten (leer = Kaufpreis × GrESt)", "neu_Nebenkosten",
         FMT_EURO, False, minimum=0),
    Feld("afa_satz", "AfA-Satz", "neu_AfASatz", FMT_PROZENT, True, 10, minimum=0, maximum=0.2),
    Feld("mietrendite", "Mietrendite auf Kaufpreis", "neu_Mietrendite", FMT_PROZENT, True, 10,
         minimum=0, maximum=0.2),
    Feld("erhaltungsquote", "Erhaltung in % vom Kaufpreis", "neu_ErhQuote", FMT_PROZENT, True,
         10, minimum=0, maximum=0.1),
    Feld("quelle", "Quelle RücklageID (leer = ohne Übertrag)", "neu_Quelle", FMT_TEXT, False,
         16, auswahl_bereich="rl_RuecklageID"),
]
NEU_STATUS_NAME = "neu_Status"

NEU_SPALTEN = [
    Rechenspalte("ak_gesamt", "AK gesamt mit Nebenkosten", "neu_AKGesamt", FMT_EURO, 16),
    Rechenspalte("ak_geb", "AK Gebäude brutto", "neu_AKGeb", FMT_EURO, 16),
    Rechenspalte("ak_gub", "AK G+B brutto", "neu_AKGuB", FMT_EURO, 16),
    Rechenspalte("ueb_geb", "aus Gebäude-Rücklage auf Gebäude", "neu_UebGeb", FMT_EURO, 16),
    Rechenspalte("ueb_gub_gub", "aus G+B-Rücklage auf G+B", "neu_UebGuBGuB", FMT_EURO, 16),
    Rechenspalte("ueb_gub_geb", "aus G+B-Rücklage auf Gebäude", "neu_UebGuBGeb", FMT_EURO, 16),
    Rechenspalte("uebertrag", "Übertrag gesamt", "neu_Uebertrag", FMT_EURO, 16),
    Rechenspalte("afa_basis", "AfA-Basis Gebäude", "neu_AfABasis", FMT_EURO, 16),
    Rechenspalte("bw_gub", "Buchwert G+B", "neu_BuchwertGuB", FMT_EURO, 16),
]


# Blatt Liquidität: je Prognosejahr eine Zeile, darunter die Summe über alle Jahre
LIQUIDITAET_SPALTEN = [
    Rechenspalte("jahr", "Jahr", "liq_Jahr", FMT_JAHR, 8),
    Rechenspalte("miete", "Miete", "liq_Miete", FMT_EURO, summe=True),
    Rechenspalte("erhaltung", "Erhaltung", "liq_Erhaltung", FMT_EURO, summe=True),
    Rechenspalte("afa", "AfA (nicht zahlungswirksam)", "liq_AfA", FMT_EURO, summe=True),
    Rechenspalte("ergebnis", "laufendes Ergebnis", "liq_Ergebnis", FMT_EURO, 16, summe=True),
    Rechenspalte("ueberschuss", "laufender Überschuss (Miete − Erhaltung)", "liq_Ueberschuss",
                 FMT_EURO, 16, summe=True),
    Rechenspalte("erloes", "Verkaufserlös netto", "liq_Erloes", FMT_EURO, 16, summe=True),
    Rechenspalte("bw_rueckfluss", "davon Buchwert-Rückfluss", "liq_BuchwertRueckfluss",
                 FMT_EURO, 16, summe=True),
    Rechenspalte("gewinn", "davon Veräußerungsgewinn", "liq_Gewinn", FMT_EURO, 16, summe=True),
    Rechenspalte("steuer_laufend", "Steuer auf laufendes Ergebnis", "liq_SteuerLaufend",
                 FMT_EURO, 16, summe=True),
    Rechenspalte("steuer_verkauf", "Steuer auf Veräußerung und Auflösung", "liq_SteuerVerkauf",
                 FMT_EURO, 16, summe=True),
    Rechenspalte("steuer", "Steuer gesamt", "liq_Steuer", FMT_EURO, summe=True),
    Rechenspalte("reinvest", "Kauf Neuobjekte inkl. Nebenkosten", "liq_Reinvest", FMT_EURO, 16,
                 summe=True),
    Rechenspalte("mittelzufluss", "freier Mittelzufluss", "liq_Mittelzufluss", FMT_EURO, 16,
                 summe=True),
    Rechenspalte("mittelzufluss_kum", "freier Mittelzufluss kumuliert", "liq_MittelzuflussKum",
                 FMT_EURO, 16),
]

# Blatt Auswertung: Kennzahlen je Prognosejahr, darunter die Summe
AUSWERTUNG_SPALTEN = [
    Rechenspalte("jahr", "Jahr", "aw_Jahr", FMT_JAHR, 8),
    Rechenspalte("ergebnis", "laufendes Ergebnis", "aw_Ergebnis", FMT_EURO, 16, summe=True),
    Rechenspalte("steuerpflichtig_vk", "steuerpflichtig aus Verkauf und Auflösung",
                 "aw_SteuerpflichtigVerkauf", FMT_EURO, 18, summe=True),
    Rechenspalte("guv", "Gesamt-GuV vor Steuern", "aw_GuV", FMT_EURO, 16, summe=True),
    Rechenspalte("steuer", "Steuer", "aw_Steuer", FMT_EURO, summe=True),
    Rechenspalte("nach_steuer", "Ergebnis nach Steuer", "aw_NachSteuer", FMT_EURO, 16,
                 summe=True),
    Rechenspalte("steuer_kum", "Steuer kumuliert", "aw_SteuerKum", FMT_EURO, 16),
    Rechenspalte("buchwert", "Buchwert Immobilien Jahresende", "aw_Buchwert", FMT_EURO, 16),
    Rechenspalte("verkehrswert", "Verkehrswert Immobilien Jahresende", "aw_Verkehrswert",
                 FMT_EURO, 16),
    Rechenspalte("stille_reserven", "stille Reserven", "aw_StilleReserven", FMT_EURO, 16),
    Rechenspalte("ruecklage", "§ 6b-Rücklage Stand", "aw_Ruecklage", FMT_EURO, 16),
    Rechenspalte("mittel_kum", "freier Mittelzufluss kumuliert", "aw_MittelzuflussKum",
                 FMT_EURO, 16),
    Rechenspalte("zins", "Zinsertrag Alternativanlage", "aw_Zins", FMT_EURO, 16, summe=True),
    Rechenspalte("steuer_zins", "Steuer auf Zinsertrag", "aw_SteuerZins", FMT_EURO, 16,
                 summe=True),
    Rechenspalte("anlage", "Alternativanlage (freie Mittel) Stand Jahresende", "aw_Anlage",
                 FMT_EURO, 18),
    Rechenspalte("latente_steuer", "latente Steuer auf stille Reserven und Rücklage",
                 "aw_LatenteSteuer", FMT_EURO, 18),
    Rechenspalte("vermoegen", "Vermögen nach Steuern", "aw_Vermoegen", FMT_EURO, 16),
]


@dataclass(frozen=True)
class Kennzahl:
    """Eine Zeile im Blatt Vergleich."""
    key: str
    bezeichnung: str
    erlaeuterung: str
    fett: bool = False


# Blatt Vergleich: Kennzahlen über den ganzen Prognosezeitraum, je eine Zeile
VERGLEICH_KENNZAHLEN = [
    Kennzahl("verkehrswert", "Verkehrswert Immobilien Ende", "Bestand am Ende des letzten Jahres"),
    Kennzahl("anlage", "Alternativanlage Ende", "freie Mittel samt Zinsen nach Steuern"),
    Kennzahl("latente_steuer", "latente Steuer Ende",
             "(stille Reserven + Rücklage) × Grenzsteuersatz, fällig erst bei Verkauf "
             "bzw. Auflösung"),
    Kennzahl("vermoegen", "Endvermögen nach Steuern",
             "Verkehrswert + Alternativanlage − latente Steuer", fett=True),
    Kennzahl("buchwert", "Buchwert Immobilien Ende", ""),
    Kennzahl("stille_reserven", "stille Reserven Ende", "Verkehrswert − Buchwert"),
    Kennzahl("ruecklage", "§ 6b-Rücklage Ende", "noch nicht übertragen oder aufgelöst"),
    Kennzahl("miete", "Miete gesamt", "A: inklusive Neuobjekt"),
    Kennzahl("afa", "AfA gesamt", "A: geringer durch den Übertrag (verlorene AfA); C: volle AfA"),
    Kennzahl("ergebnis", "laufendes Ergebnis gesamt", ""),
    Kennzahl("steuer", "Steuer auf Ergebnis, Verkauf und Auflösung",
             "B, C: Steuer auf den Gewinn sofort; A: gestundet"),
    Kennzahl("zins", "Zinsertrag Alternativanlage gesamt", "Alternativrendite vor Steuern"),
    Kennzahl("steuer_zins", "Steuer auf Zinsertrag gesamt", ""),
    Kennzahl("steuer_gesamt", "Steuer gesamt gezahlt", "ohne latente Steuer"),
    Kennzahl("reinvest", "Kauf Neuobjekte gesamt", "inklusive Nebenkosten"),
]


@dataclass
class Neuobjekt:
    """Ein Reinvestitionsobjekt. None = Feld leer lassen."""
    neu_id: Optional[str]
    name: Optional[str] = None
    kaufjahr: Optional[int] = None
    kaufpreis: Optional[float] = None
    anteil_gub: Optional[float] = None
    nebenkosten: Optional[float] = None
    afa_satz: Optional[float] = None
    mietrendite: Optional[float] = None
    erhaltungsquote: Optional[float] = None
    quelle: Optional[str] = None


@dataclass
class Verkauf:
    """Ein geplanter Verkauf. None = Feld leer lassen."""
    objekt_id: Optional[str]
    jahr: Optional[int] = None
    preis: Optional[float] = None
    faktor: Optional[float] = None
    kosten: Optional[float] = None
    aufteilung: Optional[str] = None
    nutzung_6b: Optional[str] = None


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
    # Abweichende Parameterwerte, z. B. {"par_Steuerwelt": "..."}
    parameter: dict = field(default_factory=dict)
    verkaeufe: list = field(default_factory=list)
    neuobjekte: list = field(default_factory=list)
    # gespeicherte Szenarioergebnisse im Blatt Vergleich: {"A": {key: Wert}, "B": ..., "C": ...}
    vergleich: dict = field(default_factory=dict)
