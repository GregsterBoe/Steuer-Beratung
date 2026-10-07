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
    abschnitt: Optional[str] = None  # Zwischenüberschrift vor diesem Parameter


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
    Parameter("par_Alternativrendite", "Rendite Alternativanlage p. a.", 0.03, FMT_PROZENT,
              "Zins auf die Liquidität des Vorjahresendes in allen Szenarien, voll steuerpflichtig; "
              "negative Liquidität kostet denselben Satz. Platzhalter"),
    Parameter("par_6bVorbesitz", "§ 6b Mindest-Vorbesitzzeit (Jahre)", 6, FMT_ZAHL, ""),
    Parameter("par_6bFrist", "§ 6b Reinvestitionsfrist (Jahre)", 4, FMT_ZAHL, ""),
    Parameter("par_6bFristNeubau", "§ 6b Frist bei Neubau (Jahre)", 6, FMT_ZAHL,
              "wenn mit dem Neubau vor Ende der Regelfrist begonnen wurde (§ 6b Abs. 3)"),
    Parameter("par_6bZuschlag", "§ 6b Gewinnzuschlag je Jahr", 0.06, FMT_PROZENT,
              "bei Auflösung ohne Reinvestition"),
    Parameter("par_DOGrenze", "Drei-Objekt-Grenze (Verkäufe)", 3, FMT_ZAHL,
              "mehr Verkäufe im Zeitraum: Warnung gewerblicher Grundstückshandel"),
    Parameter("par_DOJahre", "Zeitraum Drei-Objekt-Grenze (Jahre)", 5, FMT_ZAHL,
              "vereinfacht: Verkäufe innerhalb dieses Zeitraums; fachlich prüfen"),
    Parameter("par_StatusPruefung", "Plausibilitätsprüfung", "=pr_Gesamt", FMT_TEXT,
              "berechnet; Einzelheiten im Blatt Prüfung"),
    Parameter("par_AnlStand", "Stand Anlagenverzeichnis (Wj-Ende)", "=par_Basisjahr-1", FMT_JAHR,
              "Jahr der Buchwerte im Blatt Anlagen; die AfA bis zum Ende des Basisjahrs "
              "rechnet das Blatt fort"),
    # Auffülllogik: stehen als Formel (blau) in leeren Eingabezellen, überschreibbar
    Parameter("par_AnnVervielfaeltiger", "Verkehrswert = Jahresmiete ×", 20, "0.0",
              "Vervielfältiger, 20 = Bruttomietrendite 5 %; Platzhalter",
              abschnitt="Annahmen bei fehlenden Daten (blau in den Eingabeblättern)"),
    Parameter("par_AnnGebaeudeanteil", "Gebäudeanteil am Wert", 0.75, FMT_PROZENT,
              "Anteil Gebäude an AK und Verkehrswert, Rest G+B"),
    Parameter("par_AnnAfASatz", "AfA-Satz Bestand", 0.02, FMT_PROZENT,
              "Wohngebäude, Fertigstellung ab 1925 bis 2022: 2 %"),
    Parameter("par_AnnHaltedauer", "Jahre seit Kauf", 15, FMT_ZAHL,
              "Kaufjahr = Basisjahr − Jahre seit Kauf"),
    Parameter("par_AnnErhQuote", "Erhaltung in % der Miete", 0.10, FMT_PROZENT,
              "nur ohne Erhaltung aus der Buchhaltung"),
    Parameter("par_AnnGebaeudealter", "Gebäudealter im Basisjahr (Jahre)", 40, FMT_ZAHL,
              "Baujahr = Basisjahr − Gebäudealter, solange kein Baujahr erfasst ist"),
    Parameter("par_AnnReinvestJahre", "Kauf nach Verkauf (Jahre)", 1, FMT_ZAHL,
              "Kaufjahr = Verkaufsjahr + Jahre; innerhalb der § 6b-Frist halten",
              abschnitt="Annahmen Reinvestition (Verkauf mit „reinvestieren = ja“)"),
    Parameter("par_AnnReinvestQuote", "reinvestiert in % des Nettoerlöses", 1.0, FMT_PROZENT,
              "inklusive Kaufnebenkosten"),
    Parameter("par_AnnNeuNebenkosten", "Kaufnebenkosten in % des Kaufpreises", 0.07,
              FMT_PROZENT, "Grunderwerbsteuer, Notar, Grundbuch; Platzhalter"),
    Parameter("par_AnnNeuAnteilGuB", "Anteil G+B Neuobjekt", 0.25, FMT_PROZENT, ""),
    Parameter("par_AnnNeuAfASatz", "AfA-Satz Neuobjekt", 0.03, FMT_PROZENT,
              "Wohngebäude, Fertigstellung ab 2023: 3 %"),
    Parameter("par_AnnNeuAfAMethode", "AfA-Methode Neuobjekt", "linear", FMT_TEXT,
              "degressiv nur für Neubau mit Baubeginn 10/2023 bis 9/2029",
              ("linear", "degressiv")),
    Parameter("par_AnnNeuMietrendite", "Mietrendite Neuobjekt", 0.045, FMT_PROZENT,
              "Jahresmiete in % des Kaufpreises; Platzhalter"),
    Parameter("par_AnnNeuErhQuote", "Erhaltung Neuobjekt in % des Kaufpreises", 0.005,
              FMT_PROZENT, "Neubau etwa 0,3–0,5 %; in den ersten Jahren gilt zusätzlich der "
              "Anlauffaktor unten"),
    # Erhaltung nach Gebäudealter: gilt für alle Objekte, auch mit Werten aus der Buchhaltung
    Parameter("par_ErhAlterungAb", "Alterung ab Gebäudealter (Jahre)", 30, FMT_ZAHL,
              "ab diesem Alter steigt die Erhaltung zusätzlich zur Erhaltungssteigerung",
              abschnitt="Erhaltung nach Gebäudealter und Großmaßnahmen"),
    Parameter("par_ErhAlterung", "zusätzliche Steigerung je Jahr über dem Alter", 0.015,
              FMT_PROZENT, "0 % = keine Alterung; Platzhalter"),
    Parameter("par_NeuErhAnlaufJahre", "Neuobjekt: Anlaufjahre nach Kauf", 10, FMT_ZAHL,
              "so viele Jahre nach dem Kauf ist die Erhaltung gemindert"),
    Parameter("par_NeuErhAnlaufFaktor", "Neuobjekt: Erhaltung in den Anlaufjahren", 0.5,
              FMT_PROZENT, "100 % = keine Minderung; unterstellt Neubau oder frisch saniert"),
    Parameter("par_SanAlter", "Großmaßnahme fällig bei Gebäudealter", 50, FMT_ZAHL,
              "Dach, Heizung, Fassade, energetische Sanierung; Annahme nur ohne Eingabe "
              "im Objektblatt"),
    Parameter("par_SanVorlauf", "Großmaßnahme frühestens nach (Jahren)", 2, FMT_ZAHL,
              "bei schon älteren Gebäuden: frühestens erstes Prognosejahr + Vorlauf; ein "
              "Verkauf davor erspart sie"),
    Parameter("par_SanQuote", "Großmaßnahme in % des Gebäudewerts", 0.15, FMT_PROZENT,
              "Gebäudewert = Verkehrswert × Gebäudeanteil, in heutigen Preisen; 0 % = keine "
              "Großmaßnahmen annehmen"),
]

# Codenamen für VBA: ASCII, unabhängig vom angezeigten Blattnamen
CODENAME_MAPPE = "ThisWorkbook"
CODENAMEN = {
    "Start": "wsStart", "Übersicht": "wsUebersicht", "Vergleich": "wsVergleich", "Parameter": "wsParameter",
    "Objekte": "wsObjekte", "Verkäufe": "wsVerkaeufe", "Neuobjekte": "wsNeuobjekte",
    "Prognose": "wsPrognose", "Rücklagen": "wsRuecklagen", "Liquidität": "wsLiquiditaet",
    "Auswertung": "wsAuswertung", "Prüfung": "wsPruefung", "Varianten": "wsVarianten",
    "BWA-Zuordnung": "wsBWAZuordnung", "Anlagen": "wsAnlagen", "AfA-Plan": "wsAfAPlan",
    "Neukauf-KSt": "wsNeukaufKSt",
}


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
    annahme: bool = False   # leer gelassen steht eine Annahmeformel in der Zelle (blau)
    kritisch: bool = False  # Annahme ist kritisch, sobald das Objekt verkauft wird (orange)
    hinweis: str = ""       # Eingabehilfe beim Anklicken der Zelle
    # Formel leitet nur aus anderen Feldern ab: blau, zählt aber nicht als eigene Annahme
    abgeleitet: bool = False


OBJEKT_FELDER = [
    # laufende Werte, meist aus der Buchhaltung (BWA)
    Feld("objekt_id", "ObjektID", "obj_ID", FMT_TEXT, True, 12,
         hinweis="Kostenstelle bzw. eindeutige Kennung, z. B. KSt 1"),
    Feld("name", "Objektname", "obj_Name", FMT_TEXT, False, 28),
    Feld("miete", "Miete Basisjahr", "obj_MieteBasis", FMT_EURO, True, minimum=0,
         hinweis="Pflicht: Jahresmiete laut Buchhaltung (BWA 1020)"),
    Feld("erhaltung", "Erhaltung Basisjahr", "obj_ErhBasis", FMT_EURO, False, minimum=0,
         annahme=True, hinweis="BWA 1250; leer: Miete × Erhaltungsquote (Parameter)"),
    Feld("weitere_einnahmen", "weitere Einnahmen Basisjahr", "obj_EinnBasis", FMT_EURO, False,
         minimum=0, hinweis="BWA 1090; leer = 0"),
    Feld("weitere_ausgaben", "weitere Ausgaben Basisjahr", "obj_AusgBasis", FMT_EURO, False,
         minimum=0, hinweis="BWA 1100–1220, 1260; leer = 0"),
    Feld("afa_bwa", "AfA Basisjahr lt. Buchhaltung", "obj_AfABWA", FMT_EURO, False, minimum=0,
         hinweis="BWA 1240; 0 = keine AfA mehr. Daraus die Annahmen AfA je Jahr und "
                 "AK Gebäude = AfA / AfA-Satz"),
    # steuerliche Stammdaten (Anlagenverzeichnis); fehlen sie, greift die Annahme
    Feld("verkehrswert", "Verkehrswert aktuell", "obj_Verkehrswert", FMT_EURO, False, minimum=0,
         annahme=True, hinweis="leer: Jahresmiete × Vervielfältiger (Parameter)"),
    Feld("vk_quote_gebaeude", "Verkehrswertanteil Gebäude", "obj_VKQuoteGeb", FMT_PROZENT, False,
         minimum=0, maximum=1, annahme=True, kritisch=True,
         hinweis="leer: Gebäudeanteil (Parameter); teilt den Verkaufserlös auf"),
    Feld("afa_satz", "AfA-Satz", "obj_AfASatz", FMT_PROZENT, False, 10, minimum=0, maximum=0.2,
         annahme=True, kritisch=True, hinweis="leer: AfA-Satz Bestand (Parameter)"),
    Feld("kaufjahr", "Kaufjahr", "obj_Kaufjahr", FMT_JAHR, False, 10,
         minimum=1900, maximum=2100, ganzzahl=True, annahme=True, kritisch=True,
         hinweis="leer: frühester Zugang G+B oder Gebäude im Blatt Anlagen, sonst Basisjahr − "
                 "Jahre seit Kauf (Parameter); zählt für die § 6b-Vorbesitzzeit"),
    Feld("ak_gebaeude", "AK Gebäude", "obj_AKGebaeude", FMT_EURO, False, minimum=0,
         annahme=True, kritisch=True,
         hinweis="leer: Summe AHK der abnutzbaren Anlagen (Blatt Anlagen), sonst AfA lt. "
                 "Buchhaltung / AfA-Satz, sonst Verkehrswert × Gebäudeanteil, abgezinst bis "
                 "zum Kaufjahr"),
    Feld("ak_gub", "AK G+B", "obj_AKGuB", FMT_EURO, False, minimum=0, annahme=True,
         kritisch=True, hinweis="leer: Buchwert G+B im Blatt Anlagen, sonst AK Gebäude × "
                                "G+B-Anteil / Gebäudeanteil"),
    Feld("restbuchwert", "Restbuchwert Gebäude Basisjahr", "obj_Restbuchwert", FMT_EURO, False,
         minimum=0, annahme=True, kritisch=True,
         hinweis="leer: Buchwert der abnutzbaren Anlagen Ende Basisjahr (Blatt Anlagen), sonst "
                 "AK Gebäude − AfA je Jahr seit Kauf; 0, wenn die Buchhaltung keine AfA mehr zeigt"),
    # die AfA der Buchhaltung läuft in der Prognose weiter, bis der Restbuchwert verbraucht ist
    Feld("afa_jahr", "AfA je Jahr (Prognose)", "obj_AfAJahr", FMT_EURO, False, minimum=0,
         annahme=True, abgeleitet=True,
         hinweis="leer: AfA im ersten Prognosejahr lt. Blatt Anlagen (die Prognose folgt dann "
                 "je Anlage deren Buchwert), sonst AfA lt. Buchhaltung (auch 0), sonst AK Gebäude "
                 "× AfA-Satz; läuft bis der Restbuchwert verbraucht ist"),
    # Erhaltung nach Alter; die Großmaßnahme zählt als sofort abziehbarer Erhaltungsaufwand
    Feld("baujahr", "Baujahr", "obj_Baujahr", FMT_JAHR, False, 10, minimum=1800, maximum=2100,
         ganzzahl=True, annahme=True,
         hinweis="leer: frühestes AHK-Datum der Gebäude im Blatt Anlagen (orange: stimmt nur "
                 "bei Neubau, bei gekauftem Bestandsgebäude echtes Baujahr eintippen), sonst "
                 "Basisjahr − Gebäudealter (Parameter); steuert Alterung und Großmaßnahme"),
    Feld("san_jahr", "Großmaßnahme Jahr", "obj_SanJahr", FMT_JAHR, False, 11, minimum=0,
         maximum=2100, ganzzahl=True, annahme=True,
         hinweis="0 = keine; leer: Baujahr + Alter für Großmaßnahme (Parameter), "
                 "frühestens erstes Prognosejahr + Vorlauf"),
    Feld("san_betrag", "Großmaßnahme Betrag (heutige Preise)", "obj_SanBetrag", FMT_EURO, False,
         minimum=0, annahme=True,
         hinweis="leer: Verkehrswert × Gebäudeanteil × Quote (Parameter); "
                 "wächst mit der Erhaltungssteigerung bis zum Jahr der Maßnahme"),
]
# Spalten, die beim Einlesen aus der Buchhaltung kommen (grün, solange unverändert)
OBJEKT_EINGELESEN = ("name", "miete", "erhaltung", "weitere_einnahmen", "weitere_ausgaben",
                     "afa_bwa")
# davon mit Kostenstellenblatt verknüpft: Formel auf die Spalte Basisjahr des BWA-Blatts
KST_VERKNUEPFT = ("miete", "erhaltung", "weitere_einnahmen", "weitere_ausgaben", "afa_bwa")

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


# Blatt Anlagen (Anlagenverzeichnis, Projektplan Abschnitt 21): eine Zeile je Anlagegut,
# zugeordnet über die ObjektID. Liefert AK, Buchwert, Kaufjahr und die AfA je Jahr.
MIN_ANLAGEN = 300          # Zeilen mindestens; mehr, wenn mehr eingelesen werden
ART_GUB, ART_GEBAEUDE, ART_BGA, ART_IM_BAU = "G+B", "Gebäude", "BGA", "im Bau"
ART_FINANZ, ART_SONSTIGE = "Finanzanlage", "sonstige"
ANLAGE_ARTEN = (ART_GUB, ART_GEBAEUDE, ART_BGA, ART_IM_BAU, ART_FINANZ, ART_SONSTIGE)
# gehören zum Gebäudebuchwert des Objekts (abnutzbar, im Bau noch ohne AfA)
ARTEN_ABNUTZBAR = (ART_GEBAEUDE, ART_BGA, ART_IM_BAU)
GRUPPE_ABNUTZBAR = "abnutzbar"
METHODE_KEINE = "keine"
AFA_LINEAR = "linear"
AFA_DEGRESSIV = "degressiv"
ANLAGE_METHODEN = (AFA_LINEAR, AFA_DEGRESSIV, METHODE_KEINE)
ANLAGE_FELDER = [
    Feld("nr", "Inventar-Nr.", "anl_Nr", FMT_TEXT, True, 14,
         hinweis="eindeutig je Anlage; bei Gruppen aus dem Kostenstellenblatt Kostenstelle "
                 "und Bezeichnung"),
    Feld("bezeichnung", "Bezeichnung", "anl_Bez", FMT_TEXT, False, 24),
    Feld("konto", "Konto", "anl_Konto", FMT_ZAHL, False, 8),
    Feld("kost1", "KOST1", "anl_KOST1", FMT_TEXT, False, 8,
         hinweis="Kostenstelle lt. Anlagenverzeichnis; nur zur Information"),
    Feld("objekt_id", "ObjektID", "anl_ID", FMT_TEXT, False, 12,
         hinweis="ObjektID im Blatt Objekte; beim Einlesen aus KOST1 (1 = KSt 1), ohne KOST1 "
                 "über die Bezeichnung (Spalte Zuordnung). Leer: die Anlage zählt zu keinem "
                 "Objekt"),
    Feld("zuordnung", "Zuordnung", "anl_Zuordnung", FMT_TEXT, False, 26,
         hinweis="woher die ObjektID kommt: KOST1, „Bezeichnung …“ (Treffer über die "
                 "Bezeichnung, orange: bitte prüfen, danach „geprüft“ eintragen) oder "
                 "„mehrdeutig …“ (mehrere Objekte passen, ObjektID von Hand)"),
    Feld("art", "Art", "anl_Art", FMT_TEXT, True, 12, auswahl=ANLAGE_ARTEN,
         hinweis="G+B: Grund und Boden, keine AfA. Gebäude, BGA, im Bau: abnutzbar, bilden den "
                 "Gebäudebuchwert. Finanzanlage, sonstige: nicht im Modell"),
    Feld("datum", "AHK-Datum", "anl_Datum", "DD.MM.YYYY", False, 11,
         hinweis="Zugang; der früheste Zugang G+B oder Gebäude ist das Kaufjahr des Objekts"),
    Feld("ahk", "AHK", "anl_AHK", FMT_EURO, False, minimum=0),
    Feld("bw_stand", "Buchwert Stand (Wj-Ende)", "anl_BWStand", FMT_EURO, True, minimum=0,
         hinweis="Buchwert am Ende des Jahres Stand Anlagenverzeichnis (Parameterblatt)"),
    Feld("afa_art", "AfA-Art lt. Inventar", "anl_AfAArt", FMT_TEXT, False, 12),
    Feld("methode", "AfA-Methode", "anl_Methode", FMT_TEXT, True, 11, auswahl=ANLAGE_METHODEN,
         hinweis="linear: AfA p. a. bis der Buchwert verbraucht ist; degressiv: AfA-Satz vom "
                 "Buchwert; keine: G+B, im Bau, Finanzanlagen"),
    Feld("satz", "AfA-Satz", "anl_Satz", FMT_PROZENT, False, 9, minimum=0, maximum=1),
    Feld("afa", "AfA p. a.", "anl_AfA", FMT_EURO, False, minimum=0, annahme=True,
         abgeleitet=True,
         hinweis="leer: linear AHK × AfA-Satz, auf volle Euro aufgerundet wie DATEV; degressiv "
                 "Buchwert Stand × Satz. Eintippen ersetzt die Formel"),
    Feld("afa_stand", "AfA im Stand-Jahr lt. Inventar", "anl_AfAStand", FMT_EURO, False,
         hinweis="nur zur Kontrolle; enthält zeitanteilige AfA bei Zugang im Jahr"),
]
ANLAGE_SPALTEN = [
    Spalte("gruppe", "Gruppe", "anl_Gruppe", FMT_TEXT, 11),
    Spalte("zugang", "Zugangsjahr G+B/Gebäude", "anl_Zugang", FMT_JAHR, 10),
    Spalte("bw_basis", "Buchwert Ende Basisjahr", "anl_BWBasis", FMT_EURO, 15),
    Spalte("afa_basis", "AfA Basisjahr", "anl_AfABasis", FMT_EURO),
]
ANLAGE_STATUS_NAME = "anl_Status"
# berechnete Spalten im Blatt Objekte nach dem Status: was das Blatt Anlagen je Objekt liefert
OBJEKT_ANLAGEN_SPALTEN = [
    Spalte("anl_abn", "Anlagen abnutzbar", "obj_AnlAbn", FMT_ZAHL, 10),
    Spalte("anl_ak", "davon mit AHK", "obj_AnlAK", FMT_ZAHL, 9),
    Spalte("anl_gub", "Anlagen G+B", "obj_AnlGuB", FMT_ZAHL, 9),
    Spalte("anl_kauf", "Anlagen mit Zugangsjahr", "obj_AnlKauf", FMT_ZAHL, 10),
    Spalte("anl_bau", "Baujahr lt. Anlagen (frühester Gebäudezugang)", "obj_AnlBau", FMT_JAHR,
           12),
    Spalte("anl_bau_offen", "Baujahr aus Anlagen ungeprüft", "obj_AnlBauOffen", FMT_ZAHL, 10),
    Spalte("anl_afa", "AfA Basisjahr lt. Anlagen", "obj_AnlAfA", FMT_EURO, 14),
    Spalte("anl_diff", "Abweichung zur AfA lt. Buchhaltung", "obj_AnlDiff", FMT_EURO, 14),
]
# Spalten nach den Eingabefeldern: Status, Annahmen, kritische Annahmen, dann die obigen
OBJEKT_ERSTE_ANLAGEN_SPALTE = 3
ANLAGE_JAHRE_NAME = "anl_AfAJahre"   # AfA je Prognosejahr, eine Spalte je Jahr
# Statustexte; nur OK zählt in die Objekte
ANL_NICHT_IM_MODELL = "nicht im Modell (Art)"
ANL_OHNE_OBJEKT = "ohne ObjektID"
# Herkunft der ObjektID im Blatt Anlagen (Spalte Zuordnung)
ZUORDNUNG_KOST1 = "KOST1"
ZUORDNUNG_BEZEICHNUNG = "Bezeichnung"   # Treffer über die Bezeichnung, prüfen
ZUORDNUNG_MEHRDEUTIG = "mehrdeutig"
ANL_OBJEKT_FEHLT = "ObjektID fehlt im Blatt Objekte"


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
    # Etappe 8: Szenarien B und C ohne § 6b-Übertragung. 1 = Neuobjekt mit Quelle-Rücklage
    # (entfällt in Szenario B); für Bestandsobjekte gleich den Werten oben
    Spalte("mit_quelle", "Neuobjekt mit Rücklage", "prg_MitQuelle", FMT_ZAHL, 10),
    Spalte("afa_ohne6b", "AfA ohne § 6b", "prg_AfAOhne6b", FMT_EURO),
    Spalte("buchwert_ohne6b", "Buchwert Gebäude ohne § 6b", "prg_BuchwertOhne6b", FMT_EURO, 16),
    Spalte("buchwert_gub_ohne6b", "Buchwert G+B ohne § 6b", "prg_BuchwertGuBOhne6b", FMT_EURO, 16),
    # Erhaltung, als würde das Objekt nie verkauft (Baseline, Ausgangsfall im Sonderbereich)
    Spalte("erhaltung_halten", "Erhaltung bei Halten", "prg_ErhaltungHalten", FMT_EURO),
    # Sonderbereich Verkauf und Kauf: RücklageID, aus der das Neuobjekt gekauft ist
    Spalte("quelle", "Quelle RücklageID", "prg_Quelle", FMT_TEXT, 14),
    # AfA, als würde das Objekt nie verkauft: je Anlage aus dem Blatt Anlagen, sonst AfA je
    # Jahr bis zum Restbuchwert; der Plan nimmt sie bis zum Verkaufsjahr
    Spalte("afa_halten", "AfA bei Halten", "prg_AfAHalten", FMT_EURO),
]

# Blatt Verkäufe (Etappe 4, Projektplan Abschnitt 11): Eingaben, dann berechnete Spalten
MAX_VERKAEUFE = 50
VERKAUF_FELDER = [
    Feld("objekt_id", "ObjektID", "vk_ID", FMT_TEXT, True, 14),
    Feld("jahr", "Verkaufsjahr", "vk_Jahr", FMT_JAHR, True, 12,
         minimum=1900, maximum=2100, ganzzahl=True),
    Feld("preis", "Verkaufspreis", "vk_Preis", FMT_EURO, True, minimum=0, annahme=True,
         kritisch=True, hinweis="leer: Verkehrswert, fortgeschrieben bis zum Verkaufsjahr"),
    Feld("kosten", "Verkaufskosten", "vk_Kosten", FMT_EURO, False, minimum=0),
    Feld("anteil_gub", "Anteil G+B lt. Kaufvertrag", "vk_AnteilGuBVertrag", FMT_PROZENT, False,
         minimum=0, maximum=1),
    Feld("nutzung_6b", "§ 6b nutzen", "vk_6b", FMT_TEXT, False, 10, auswahl=("ja", "nein")),
    # verlängert die Reinvestitionsfrist auf par_6bFristNeubau
    Feld("neubau_6b", "§ 6b Neubau begonnen", "vk_6bNeubau", FMT_TEXT, False, 11,
         auswahl=("ja", "nein")),
    # ja = Neuobjekt in derselben Zeile des Blatts Neuobjekte aus den Annahmen (Reinvestition)
    Feld("reinvest", "reinvestieren", "vk_Reinvest", FMT_TEXT, False, 12,
         auswahl=("ja", "nein"),
         hinweis="ja: Kauf eines Neuobjekts nach den Annahmen auf dem Parameterblatt"),
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
STATUS_ANNAHME_GELOESCHT = "Wert fehlt: Annahme gelöscht (Makro „Annahmen wiederherstellen“)"

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
    Feld("kst", "Kostenstelle Neukauf", "ne_KSt", FMT_TEXT, False, 14,
         hinweis="z. B. KSt 31 aus dem Blatt Neukauf-KSt: dessen Jahreswerte ersetzen Miete, "
                 "weitere Einnahmen, Erhaltung und weitere Ausgaben aus Mietrendite und "
                 "Erhaltungsquote; die AfA rechnet weiter das Modell"),
]
# Blatt Neukauf-KSt: Planwerte der Neukauf-Kostenstellen (hinter „KSt 9999“) je Jahr.
# (Schlüssel, Bezeichnung, BWA-Nr., Steigerung für Jahre ohne Wert)
NEUKAUF_POSITIONEN = [
    ("miete", "Miete", "1020", "par_Mietsteig"),
    ("einnahmen", "weitere Einnahmen", "1090", "par_Mietsteig"),
    ("erhaltung", "Erhaltung", "1250", "par_Erhaltsteig"),
    ("ausgaben", "weitere Ausgaben", "1100–1220, 1260", "par_Kostensteig"),
]
MAX_NEUKAUF = 20   # Kostenstellen im Blatt Neukauf-KSt
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
    # 1 = Quelle-Rücklage angegeben; das Objekt entfällt in Szenario B
    Spalte("mit_quelle", "mit Rücklage", "ne_MitQuelle", FMT_ZAHL, 8),
]
NEU_STATUS_NAME = "ne_Status"


# Szenarien (Etappe 8, Projektplan Abschnitt 15): alle rechnen gleichzeitig, je eine Tabelle
# nebeneinander in Liquidität und Auswertung. key, Titel, Namenspräfix Liquidität, Auswertung
@dataclass(frozen=True)
class Szenario:
    key: str
    titel: str
    kurz: str
    liq: str
    aus: str


SZ_A = Szenario("A", "A Plan: § 6b-Kette wie erfasst", "A Plan (§ 6b)", "liq", "aus")
SZ_B = Szenario("B", "B: sofort versteuern, Kapital anlegen (Neuobjekte mit Rücklage entfallen)",
                "B sofort versteuern, anlegen", "lvb", "avb")
SZ_C = Szenario("C", "C: sofort versteuern, Neuobjekte trotzdem kaufen (volle AfA-Basis)",
                "C sofort versteuern, kaufen", "lvc", "avc")
SZ_BASELINE = Szenario("Baseline", "Baseline: alles halten", "Baseline halten", "lqb", "asb")
SZENARIEN = [SZ_A, SZ_B, SZ_C, SZ_BASELINE]

# Blatt Liquidität (Etappe 7 und 8, Projektplan Abschnitt 14): je Jahr eine Zeile,
# je Szenario eine Tabelle mit denselben Spalten (eine Spalte Abstand)
_LIQ = [
    ("jahr", "Jahr", "Jahr", FMT_JAHR, 8),
    ("einnahmen", "Mieten und weitere Einnahmen", "Einnahmen", FMT_EURO, 14),
    ("ausgaben", "Erhaltung und weitere Ausgaben", "Ausgaben", FMT_EURO, 14),
    ("afa", "AfA Gebäude", "AfA", FMT_EURO, 14),
    ("ergebnis", "laufendes Ergebnis", "Ergebnis", FMT_EURO, 14),
    ("verkauf", "steuerwirksam aus Verkauf und Rücklage", "Verkauf", FMT_EURO, 16),
    ("zins", "Zinsertrag Alternativanlage", "Zins", FMT_EURO, 14),
    ("zve", "Ergebnis vor Verlustvortrag", "ZvE", FMT_EURO, 14),
    ("vortrag_genutzt", "Verlustvortrag genutzt", "VortragGenutzt", FMT_EURO, 14),
    ("bemessung", "Bemessungsgrundlage", "Bemessung", FMT_EURO, 14),
    ("vortrag", "Verlustvortrag Ende", "Vortrag", FMT_EURO, 14),
    ("steuer", "Steuer", "Steuer", FMT_EURO, 14),
    ("verkaufserloes", "Verkaufserlöse netto", "Verkaufserloes", FMT_EURO, 14),
    ("rueckfluss", "davon Buchwert-Rückfluss", "Rueckfluss", FMT_EURO, 14),
    ("kauf", "Kauf Neuobjekte inkl. Nebenkosten", "Kauf", FMT_EURO, 15),
    ("zufluss", "freier Mittelzufluss", "Zufluss", FMT_EURO, 14),
    ("kum", "Liquidität kumuliert Ende", "Kum", FMT_EURO, 15),
]
# Blatt Auswertung: Kennzahlen je Jahr, gleicher Aufbau je Szenario
_AUS = [
    ("jahr", "Jahr", "Jahr", FMT_JAHR, 8),
    ("ergebnis", "laufendes Ergebnis", "Ergebnis", FMT_EURO, 14),
    ("verkauf", "steuerwirksam aus Verkauf und Rücklage", "Verkauf", FMT_EURO, 16),
    ("zins", "Zinsertrag Alternativanlage", "Zins", FMT_EURO, 14),
    ("guv", "Gesamt-GuV vor Steuern", "GuV", FMT_EURO, 14),
    ("steuer", "Steuer", "Steuer", FMT_EURO, 14),
    ("nach_steuer", "Ergebnis nach Steuern", "NachSteuer", FMT_EURO, 14),
    ("steuer_kum", "Steuer kumuliert", "SteuerKum", FMT_EURO, 14),
    ("verkehrswert", "Verkehrswert Bestand", "Verkehrswert", FMT_EURO, 15),
    ("buchwert", "Buchwert Bestand (Gebäude + G+B)", "Buchwert", FMT_EURO, 16),
    ("stille_reserven", "stille Reserven", "StilleReserven", FMT_EURO, 14),
    ("ruecklage", "§ 6b-Rücklage Bestand", "Ruecklage", FMT_EURO, 14),
    ("vortrag", "Verlustvortrag", "Vortrag", FMT_EURO, 14),
    ("liquiditaet", "Liquidität kumuliert", "Liquiditaet", FMT_EURO, 14),
    ("vermoegen", "Gesamtvermögen vor latenter Steuer", "Vermoegen", FMT_EURO, 16),
    ("latente_steuer", "latente Steuer", "LatenteSteuer", FMT_EURO, 14),
    ("vermoegen_netto", "Gesamtvermögen nach latenter Steuer", "VermoegenNetto", FMT_EURO, 17),
]


def _spalten(vorlage, praefix: str) -> list:
    return [Spalte(key, u, f"{praefix}_{name}", fmt, b) for key, u, name, fmt, b in vorlage]


def liq_spalten(sz: Szenario) -> list:
    return _spalten(_LIQ, sz.liq)


def aus_spalten(sz: Szenario) -> list:
    return _spalten(_AUS, sz.aus)


LIQ_SPALTEN = liq_spalten(SZ_A)
AUSWERTUNG_SPALTEN = aus_spalten(SZ_A)

# Blatt Vergleich (Etappe 8): Kennzahlen je Szenario. art "ende" = Wert im letzten
# Prognosejahr aus der Auswertung, "summe" = Summe über alle Jahre aus der Liquidität
VERGLEICH_KENNZAHLEN = [
    ("Endvermögen nach latenter Steuer", "aus", "VermoegenNetto", "ende",
     "Verkehrswert + Liquidität − latente Steuer am Ende des letzten Prognosejahrs"),
    ("Verkehrswert Immobilien", "aus", "Verkehrswert", "ende", "Objekte im Bestand"),
    ("Liquidität (Alternativanlage)", "aus", "Liquiditaet", "ende",
     "kumulierte freie Mittel samt Zinsen"),
    ("latente Steuer", "aus", "LatenteSteuer", "ende",
     "Steuer bei Verkauf aller Objekte und Auflösung der Rücklage"),
    ("Buchwert Immobilien", "aus", "Buchwert", "ende", "Gebäude + G+B"),
    ("stille Reserven", "aus", "StilleReserven", "ende", "Verkehrswert − Buchwert"),
    ("§ 6b-Rücklage", "aus", "Ruecklage", "ende", "noch nicht übertragen oder aufgelöst"),
    ("Verlustvortrag", "aus", "Vortrag", "ende", "mindert die latente Steuer"),
    ("Mieten und weitere Einnahmen", "liq", "Einnahmen", "summe", "Summe über alle Jahre"),
    ("Erhaltung und weitere Ausgaben", "liq", "Ausgaben", "summe", "Summe über alle Jahre"),
    ("AfA Gebäude", "liq", "AfA", "summe", "verlorene AfA durch § 6b: A gegen C"),
    ("laufendes Ergebnis", "liq", "Ergebnis", "summe", "Summe über alle Jahre"),
    ("steuerwirksam aus Verkauf und Rücklage", "liq", "Verkauf", "summe",
     "Gewinne, Auflösung und Zuschlag"),
    ("Zinsertrag Alternativanlage", "liq", "Zins", "summe", "Summe über alle Jahre"),
    ("Steuer", "liq", "Steuer", "summe", "gezahlte Steuer; die gestundete steht in der latenten"),
    ("Verkaufserlöse netto", "liq", "Verkaufserloes", "summe", "Summe über alle Jahre"),
    ("Kauf Neuobjekte", "liq", "Kauf", "summe", "Summe über alle Jahre"),
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
    afa_bwa: Optional[float] = None      # AfA im Basisjahr lt. Buchhaltung (BWA 1240)
    afa_jahr: Optional[float] = None     # AfA je Prognosejahr; leer = aus afa_bwa
    baujahr: Optional[int] = None
    san_jahr: Optional[int] = None       # Großmaßnahme; 0 = keine
    san_betrag: Optional[float] = None


@dataclass
class Anlage:
    """Eine Zeile des Anlagenverzeichnisses. None = Feld leer lassen."""
    nr: str
    bw_stand: Optional[float]
    art: Optional[str] = None
    methode: Optional[str] = None
    bezeichnung: Optional[str] = None
    konto: Optional[int] = None
    kost1: Optional[str] = None
    objekt_id: Optional[str] = None
    zuordnung: Optional[str] = None      # Herkunft der ObjektID, siehe ZUORDNUNG_*
    datum: Optional[object] = None       # datetime.date
    ahk: Optional[float] = None
    afa_art: Optional[str] = None
    satz: Optional[float] = None
    afa: Optional[float] = None          # leer = Formel aus AHK × Satz
    afa_stand: Optional[float] = None


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
    reinvest: Optional[str] = None       # ja = Neuobjekt aus den Annahmen


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
    kst: Optional[str] = None            # Neukauf-Kostenstelle, z. B. "KSt 31"


# Blatt Prüfung: Plausibilitätsprüfungen (Etappe 9, Projektplan Abschnitt 5 und 19)
FEHLER, WARNUNG, HINWEIS = "Fehler", "Warnung", "Hinweis"


@dataclass(frozen=True)
class Pruefung:
    """Eine Zeile im Blatt Prüfung; die Formel für die Anzahl liefert formeln.py."""
    key: str
    bezeichnung: str
    art: str             # Fehler, Warnung oder Hinweis
    wo: str              # wo nachsehen


PRUEFUNGEN = [
    Pruefung("objekte", "Objekte mit Status ungleich OK (Pflichtfeld, doppelte ObjektID, "
             "Kaufjahr, Restbuchwert)", FEHLER, "Blatt Objekte, Spalte Status"),
    Pruefung("verkaeufe", "Verkäufe mit Status ungleich OK (unbekanntes Objekt, Verkaufsjahr, "
             "Preis, Aufteilung); die Zeile rechnet nicht mit", FEHLER,
             "Blatt Verkäufe, Spalte Status"),
    Pruefung("neuobjekte", "Neuobjekte mit Status ungleich OK, darunter Kauf nach Fristjahr "
             "(Fristverstoß) und Kauf vor Bildung der Rücklage", FEHLER,
             "Blatt Neuobjekte, Spalte Status"),
    Pruefung("steuerwelt", "Steuerwelt außerhalb des MVP, Ergebnisse gelten nur für die GmbH",
             FEHLER, "Blatt Parameter, Steuerwelt"),
    Pruefung("vorbesitz", "§ 6b gewählt, aber Vorbesitzzeit unter der Mindestdauer: "
             "Gewinn wird sofort versteuert", WARNUNG, "Blatt Verkäufe, Spalte Status"),
    Pruefung("teiluebertrag", "Rücklage nicht voll übertragen, obwohl ein Neuobjekt sie nennt "
             "(Gebäudeanteil oder Kaufpreis zu klein): Rest wird im Fristjahr mit "
             "Gewinnzuschlag aufgelöst", WARNUNG,
             "Blatt Rücklagen, Spalte Auflösung; Blatt Neuobjekte"),
    Pruefung("ohne_reinvest", "Rücklage ohne Neuobjekt: Auflösung im Fristjahr mit "
             "Gewinnzuschlag", WARNUNG, "Blatt Rücklagen, Spalte Auflösung"),
    Pruefung("grundstueckshandel", "Verkäufe über der Drei-Objekt-Grenze im Zeitraum: Gefahr "
             "gewerblicher Grundstückshandel, Objekte wären Umlaufvermögen und § 6b entfiele",
             WARNUNG, "Blatt Verkäufe, Spalte Verkaufsjahr; Grenze auf dem Parameterblatt"),
    Pruefung("frist_ende", "Frist einer Rücklage endet nach dem Prognoseende; Auflösung und "
             "Zuschlag liegen außerhalb des Rasters", HINWEIS, "Blatt Rücklagen, Spalte Hinweis"),
    Pruefung("liquiditaet", "Liquidität im Plan (A) in mindestens einem Jahr negativ: "
             "Finanzierungsbedarf (Stufe 2)", HINWEIS, "Blatt Liquidität, Liquidität kumuliert"),
    Pruefung("kritisch", "Verkauf mit Annahmen bei steuerlichen Stammdaten oder Verkaufspreis "
             "(orange): Veräußerungsgewinn und § 6b-Rücklage sind nur geschätzt", WARNUNG,
             "Blätter Objekte und Verkäufe, orange Zellen"),
    Pruefung("bwa_zuordnung", "BWA-Zuordnung mit ungültiger BWA-Nr. oder BWA Alle Objekte "
             "weicht vom Ergebnis der Liquidität ab (Anzahl Posten bzw. Jahre)", WARNUNG,
             "Blatt BWA-Zuordnung, Spalte BWA-Nr. und Kontrolle"),
    Pruefung("annahmen", "Objekte mit Annahmen (blau): Werte aus der Auffülllogik des "
             "Parameterblatts", HINWEIS, "Blatt Objekte, Spalte Annahmen"),
    Pruefung("anlagen", "Anlagen G+B, Gebäude, BGA oder im Bau ohne Objekt (keine oder unbekannte "
             "ObjektID) oder mit fehlender Angabe: AK und Buchwert fehlen im Modell",
             HINWEIS, "Blatt Anlagen, Spalte Status"),
    Pruefung("anlagen_bez", "Anlagen ohne KOST1, die beim Einlesen über die Bezeichnung einem "
             "Objekt zugeordnet wurden (ungeprüft) oder zu mehreren Objekten passen",
             HINWEIS, "Blatt Anlagen, Spalten ObjektID und Zuordnung (orange)"),
    Pruefung("baujahr", "Objekte, deren Baujahr aus dem frühesten AHK-Datum der Gebäude stammt "
             "(orange): stimmt nur bei Neubau. Bei gekauftem Bestandsgebäude das echte "
             "Baujahr eintragen, sonst das Jahr zur Bestätigung eintippen", WARNUNG,
             "Blatt Objekte, Spalte Baujahr"),
    Pruefung("kst_abweichung", "Objekte, deren Basiswerte im Blatt Objekte von ihrem "
             "Kostenstellenblatt abweichen (orange, überschrieben): Prognose und BWA Alle Objekte "
             "passen nicht zusammen. Im Kostenstellenblatt ändern oder per Makro „Objekte → "
             "Kostenstellen“ übernehmen", WARNUNG, "Blatt Objekte, orange Zellen"),
    Pruefung("neukauf_kst", "Neuobjekte mit Kostenstelle Neukauf, die im Blatt Neukauf-KSt "
             "fehlt: das Neuobjekt rechnet mit Mietrendite und Erhaltungsquote", WARNUNG,
             "Blatt Neuobjekte, Spalte Kostenstelle Neukauf"),
    Pruefung("afa_plan", "Objekte mit AfA-Plan (Blatt AfA-Plan, meist aus BWA 1240 der "
             "Kostenstellen-Datei): in den Jahren mit Wert ersetzt er die AfA-Fortschreibung",
             HINWEIS, "Blatt AfA-Plan, Spalte Jahre mit Wert"),
    Pruefung("anlagen_afa", "Objekte, deren AfA im Basisjahr lt. Blatt Anlagen von der AfA lt. "
             "Buchhaltung (BWA 1240) um mehr als 1 € abweicht: Zuordnung der Anlagen prüfen",
             HINWEIS, "Blatt Objekte, Spalten AfA Basisjahr lt. Anlagen und lt. Buchhaltung"),
]

# Blatt Varianten: je Makrolauf eine Zeile mit festen Werten (Projektplan Abschnitt 19)
VARIANTEN_KOPF = ["Variante", "festgehalten am", "Endvermögen A", "Endvermögen B",
                  "Endvermögen C", "Endvermögen Baseline", "A − B", "A − C", "Steuer A gesamt",
                  "Prüfung"]
MAX_VARIANTEN = 100


@dataclass
class Modell:
    objekte: list = field(default_factory=list)
    verkaeufe: list = field(default_factory=list)
    neuobjekte: list = field(default_factory=list)
    # abweichende Parameterwerte, z. B. {"par_Alternativrendite": 0}
    parameter: dict = field(default_factory=dict)
    # Anlagenverzeichnis (Blatt Anlagen), je Zeile eine Anlage
    anlagen: list = field(default_factory=list)
    # Ist-Werte aus den Kostenstellenblättern je ObjektID (einlesen.LaufendeWerte);
    # ohne Eintrag zeigt das BWA-Blatt im Basisjahr die Werte des Objektblatts
    kostenstellen: dict = field(default_factory=dict)
    # Neukauf-Kostenstellen (hinter „KSt 9999“) je KSt, einlesen.LaufendeWerte mit jahre
    neukauf: dict = field(default_factory=dict)
    # Schnellcheck: § 6b und Reinvestition je Verkauf als Annahme, Details ausgeblendet
    schnellcheck: bool = False
    # Blatt BWA-Zuordnung: abweichende BWA-Nr. je Posten, z. B. {"gewinn": 1351},
    # und "verkauf": "brutto"; ohne Eintrag gilt der Standard
    bwa_zuordnung: dict = field(default_factory=dict)
