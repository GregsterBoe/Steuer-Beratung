"""Schreibt die Arbeitsmappe: Blätter, Kopfzeilen, benannte Bereiche, Validierung."""

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import formeln
from .modelle import (CODENAME_MAPPE, CODENAMEN, FEHLER, FMT_EURO, FMT_JAHR, FMT_PROZENT,
                      FMT_TEXT, FMT_ZAHL, HINWEIS, MAX_NEUOBJEKTE, MAX_VARIANTEN, PRUEFUNGEN,
                      VARIANTEN_KOPF, WARNUNG,
                      MAX_OBJEKTE, MAX_VERKAEUFE, NEU_FELDER, NEU_SPALTEN, NEU_STATUS_NAME, OBJEKT_FELDER,
                      PARAMETER, PROGNOSE_SPALTEN, RUECKLAGE_JAHR_SPALTEN, RUECKLAGE_SPALTEN,
                      STATUS_NAME, STATUS_UEBERSCHRIFT, SZ_A, SZ_BASELINE, SZENARIEN,
                      VERGLEICH_KENNZAHLEN, VERKAUF_FELDER, VERKAUF_SPALTEN, VERKAUF_STATUS_NAME, Modell,
                      aus_spalten, liq_spalten, prognosejahre)

HINWEIS_FINANZIERUNG = "Alle Werte vor Finanzierung (ohne Zins und Tilgung)."

FONT_TITEL = Font(bold=True, size=14)
FONT_KOPF = Font(bold=True, color="FFFFFF")
FILL_KOPF = PatternFill("solid", fgColor="305496")
FILL_EINGABE = PatternFill("solid", fgColor="FFF2CC")   # gelb = hier wird getippt
FILL_BERECHNET = PatternFill("solid", fgColor="E7E6E6")  # grau = Formel
FILL_FEHLER = PatternFill("solid", fgColor="F8CBAD")
FILL_WARNUNG = PatternFill("solid", fgColor="FFE699")


def _name(wb, name: str, ref: str) -> None:
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _kopf(ws, zeile: int, werte: list) -> None:
    for spalte, wert in enumerate(werte, start=1):
        c = ws.cell(row=zeile, column=spalte, value=wert)
        c.font = FONT_KOPF
        c.fill = FILL_KOPF
        c.alignment = Alignment(wrap_text=True, vertical="center")


def _validierung(ws, f, bereich: str) -> None:
    """Zahlenbereich eines Eingabefelds absichern."""
    if f.minimum is None and f.maximum is None:
        return
    dv = DataValidation(
        type="whole" if f.ganzzahl else "decimal",
        operator="between",
        formula1=str(f.minimum if f.minimum is not None else -1e15),
        formula2=str(f.maximum if f.maximum is not None else 1e15),
        showErrorMessage=True,
        errorTitle="Ungültiger Wert",
        error=f"{f.ueberschrift}: Wert außerhalb des zulässigen Bereichs.",
    )
    ws.add_data_validation(dv)
    dv.add(bereich)


def _status_rot(ws, bereich: str, erste_zelle: str) -> None:
    """Status rot, sobald er gesetzt ist und nicht OK lautet."""
    ws.conditional_formatting.add(
        bereich,
        FormulaRule(formula=[f'AND({erste_zelle}<>"",{erste_zelle}<>"OK")'], fill=FILL_FEHLER),
    )


def _blatt_parameter(wb, modell: Modell) -> None:
    ws = wb.active
    ws.title = "Parameter"
    ws["A1"] = "Parameter"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = HINWEIS_FINANZIERUNG
    _kopf(ws, 4, ["Bezeichnung", "Wert", "Name", "Erläuterung"])

    unbekannt = set(modell.parameter) - {p.name for p in PARAMETER}
    if unbekannt:
        raise ValueError(f"unbekannte Parameter: {sorted(unbekannt)}")
    for zeile, p in enumerate(PARAMETER, start=5):
        ws.cell(row=zeile, column=1, value=p.bezeichnung)
        c = ws.cell(row=zeile, column=2, value=modell.parameter.get(p.name, p.wert))
        c.number_format = p.format
        berechnet = isinstance(p.wert, str) and p.wert.startswith("=")
        c.fill = FILL_BERECHNET if berechnet else FILL_EINGABE
        ws.cell(row=zeile, column=3, value=p.name)
        ws.cell(row=zeile, column=4, value=p.erlaeuterung)
        _name(wb, p.name, f"Parameter!$B${zeile}")
        if p.auswahl:
            dv = DataValidation(type="list", formula1='"' + ",".join(p.auswahl) + '"',
                                allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(c.coordinate)

    zeile_pruefung = 5 + [p.name for p in PARAMETER].index("par_StatusPruefung")
    ws.conditional_formatting.add(
        f"B{zeile_pruefung}",
        FormulaRule(formula=[f'B{zeile_pruefung}<>"OK"'], fill=FILL_FEHLER))
    ws["F3"] = ("Steuerung per Makro (Datei als .xlsm, Makros aktivieren): die Schaltflächen "
                "unten. Gerechnet wird immer in den Formeln, auch ohne Makros.")
    ws["F3"].font = Font(italic=True)

    for spalte, breite in zip("ABCD", (36, 14, 22, 60)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A5"


def _blatt_objekte(wb, modell: Modell) -> None:
    ws = wb.create_sheet("Objekte")
    erste, letzte = 2, MAX_OBJEKTE + 1
    status_spalte = len(OBJEKT_FELDER) + 1
    _kopf(ws, 1, [f.ueberschrift for f in OBJEKT_FELDER] + [STATUS_UEBERSCHRIFT])
    ws.row_dimensions[1].height = 32

    for i, f in enumerate(OBJEKT_FELDER, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = f.breite
        _name(wb, f.name, f"Objekte!${bst}${erste}:${bst}${letzte}")
        _validierung(ws, f, f"{bst}{erste}:{bst}{letzte}")

    st = get_column_letter(status_spalte)
    ws.column_dimensions[st].width = 26
    _name(wb, STATUS_NAME, f"Objekte!${st}${erste}:${st}${letzte}")
    letzte_spalte = get_column_letter(status_spalte)
    _name(wb, "obj_Basis", f"Objekte!$A${erste}:${letzte_spalte}${letzte}")
    # nur die gelben Eingabefelder; die Makros in modObjekte schreiben nur hierhin
    _name(wb, "obj_Eingabe",
          f"Objekte!$A${erste}:${get_column_letter(len(OBJEKT_FELDER))}${letzte}")

    for zeile in range(erste, letzte + 1):
        for i, f in enumerate(OBJEKT_FELDER, start=1):
            c = ws.cell(row=zeile, column=i)
            c.number_format = f.format
            c.fill = FILL_EINGABE
        c = ws.cell(row=zeile, column=status_spalte, value=formeln.status_objekt(zeile))
        c.fill = FILL_BERECHNET

    _status_rot(ws, f"{st}{erste}:{st}{letzte}", f"{st}{erste}")

    for zeile, obj in enumerate(modell.objekte, start=erste):
        for i, f in enumerate(OBJEKT_FELDER, start=1):
            wert = getattr(obj, f.key)
            if wert is not None:
                ws.cell(row=zeile, column=i, value=wert)

    ws.freeze_panes = "B2"


def _blatt_prognose(wb) -> None:
    """Je Zeile des Objektblatts ein Block mit einer Zeile je Prognosejahr, danach je Neuobjekt."""
    ws = wb.create_sheet("Prognose")
    jahre = prognosejahre()
    erste, letzte = 2, (MAX_OBJEKTE + MAX_NEUOBJEKTE) * jahre + 1
    _kopf(ws, 1, [s.ueberschrift for s in PROGNOSE_SPALTEN])
    ws.row_dimensions[1].height = 32

    for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"Prognose!${bst}${erste}:${bst}${letzte}")

    bloecke = ([(nr, formeln.prognose_zeile) for nr in range(1, MAX_OBJEKTE + 1)]
               + [(nr, formeln.prognose_zeile_neu) for nr in range(1, MAX_NEUOBJEKTE + 1)])
    for block, (nr, zeilenformeln) in enumerate(bloecke):
        for j in range(jahre):
            zeile = erste + block * jahre + j
            formeln_zeile = zeilenformeln(zeile, nr, erstes_jahr=j == 0)
            for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
                c = ws.cell(row=zeile, column=i, value=formeln_zeile[s.key])
                c.number_format = s.format
                c.fill = FILL_BERECHNET

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(PROGNOSE_SPALTEN))}{letzte}"


def _blatt_verkaeufe(wb, modell: Modell) -> None:
    """Geplante Verkäufe: Eingaben, Aufteilung des Erlöses und Gewinn je Gebäude und G+B."""
    ws = wb.create_sheet("Verkäufe")
    erste, letzte = 2, MAX_VERKAEUFE + 1
    spalten = VERKAUF_FELDER + VERKAUF_SPALTEN
    status_spalte = len(spalten) + 1
    st = get_column_letter(status_spalte)
    _kopf(ws, 1, [s.ueberschrift for s in spalten] + [STATUS_UEBERSCHRIFT])
    ws.row_dimensions[1].height = 45

    for i, s in enumerate(spalten, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'Verkäufe'!${bst}${erste}:${bst}${letzte}")
    for i, f in enumerate(VERKAUF_FELDER, start=1):
        bst = get_column_letter(i)
        _validierung(ws, f, f"{bst}{erste}:{bst}{letzte}")
        if f.auswahl:
            dv = DataValidation(type="list", formula1='"' + ",".join(f.auswahl) + '"',
                                allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{bst}{erste}:{bst}{letzte}")
    # ObjektID als Auswahl aus dem Objektblatt
    dv = DataValidation(type="list", formula1="obj_ID", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"A{erste}:A{letzte}")

    ws.column_dimensions[st].width = 34
    _name(wb, VERKAUF_STATUS_NAME, f"'Verkäufe'!${st}${erste}:${st}${letzte}")
    for zeile in range(erste, letzte + 1):
        for i, f in enumerate(VERKAUF_FELDER, start=1):
            c = ws.cell(row=zeile, column=i)
            c.number_format = f.format
            c.fill = FILL_EINGABE
        rechnung = formeln.verkauf_zeile(zeile)
        for i, s in enumerate(VERKAUF_SPALTEN, start=len(VERKAUF_FELDER) + 1):
            c = ws.cell(row=zeile, column=i, value=rechnung[s.key])
            c.number_format = s.format
            c.fill = FILL_BERECHNET
        c = ws.cell(row=zeile, column=status_spalte, value=formeln.status_verkauf(zeile))
        c.fill = FILL_BERECHNET
    _status_rot(ws, f"{st}{erste}:{st}{letzte}", f"{st}{erste}")

    for zeile, vk in enumerate(modell.verkaeufe, start=erste):
        for i, f in enumerate(VERKAUF_FELDER, start=1):
            wert = getattr(vk, f.key)
            if wert is not None:
                ws.cell(row=zeile, column=i, value=wert)

    hinweis = get_column_letter(status_spalte + 2)
    ws[f"{hinweis}1"] = ("Verkauf zum Jahresende: Miete und AfA laufen im Verkaufsjahr noch, "
                         "ab dem Folgejahr ist das Objekt inaktiv.")
    ws[f"{hinweis}2"] = ("Aufteilung des Nettoerlöses: Anteil G+B laut Kaufvertrag, "
                         "sonst 1 − Verkehrswertanteil Gebäude aus dem Objektblatt.")
    ws[f"{hinweis}3"] = ("§ 6b Neubau begonnen = ja: Mit dem Bau wurde vor Ende der Regelfrist "
                         "begonnen, die Rücklage läuft dann par_6bFristNeubau statt par_6bFrist Jahre.")
    ws.freeze_panes = "B2"


def _blatt_neuobjekte(wb, modell: Modell) -> None:
    """Reinvestitionsobjekte: Eingaben, Übertragung der Rücklage, AfA-Basis."""
    ws = wb.create_sheet("Neuobjekte")
    erste, letzte = 2, MAX_NEUOBJEKTE + 1
    spalten = NEU_FELDER + NEU_SPALTEN
    status_spalte = len(spalten) + 1
    st = get_column_letter(status_spalte)
    _kopf(ws, 1, [s.ueberschrift for s in spalten] + [STATUS_UEBERSCHRIFT])
    ws.row_dimensions[1].height = 45

    for i, s in enumerate(spalten, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'Neuobjekte'!${bst}${erste}:${bst}${letzte}")
    for i, f in enumerate(NEU_FELDER, start=1):
        _validierung(ws, f, f"{get_column_letter(i)}{erste}:{get_column_letter(i)}{letzte}")
    # Quelle als Auswahl aus den gebildeten Rücklagen
    q = formeln.nspalte("quelle")
    dv = DataValidation(type="list", formula1="rl_ID", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{q}{erste}:{q}{letzte}")

    ws.column_dimensions[st].width = 44
    _name(wb, NEU_STATUS_NAME, f"'Neuobjekte'!${st}${erste}:${st}${letzte}")
    for zeile in range(erste, letzte + 1):
        for i, f in enumerate(NEU_FELDER, start=1):
            c = ws.cell(row=zeile, column=i)
            c.number_format = f.format
            c.fill = FILL_EINGABE
        rechnung = formeln.neu_zeile(zeile)
        for i, s in enumerate(NEU_SPALTEN, start=len(NEU_FELDER) + 1):
            c = ws.cell(row=zeile, column=i, value=rechnung[s.key])
            c.number_format = s.format
            c.fill = FILL_BERECHNET
        c = ws.cell(row=zeile, column=status_spalte, value=formeln.status_neu(zeile))
        c.fill = FILL_BERECHNET
    _status_rot(ws, f"{st}{erste}:{st}{letzte}", f"{st}{erste}")

    for zeile, neu in enumerate(modell.neuobjekte, start=erste):
        for i, f in enumerate(NEU_FELDER, start=1):
            wert = getattr(neu, f.key)
            if wert is not None:
                ws.cell(row=zeile, column=i, value=wert)

    hinweis = get_column_letter(status_spalte + 2)
    ws[f"{hinweis}1"] = ("Kauf zum Jahresende: Übertragung und Bestand im Kaufjahr, "
                         "Miete, Erhaltung und AfA ab dem Folgejahr.")
    ws[f"{hinweis}2"] = ("Übertragung: ü1 Gebäudegewinn auf Gebäude, ü2 G+B-Gewinn auf G+B, "
                         "ü3 Rest des G+B-Gewinns auf Gebäude. AfA-Basis = AK Gebäude − ü1 − ü3.")
    ws[f"{hinweis}3"] = ("Nutzen mehrere Neuobjekte dieselbe Rücklage, gilt die Zeilenreihenfolge: "
                         "jede Zeile erhält, was die Zeilen darüber übrig lassen.")
    ws[f"{hinweis}4"] = "Kaufnebenkosten werden im Verhältnis G+B zu Gebäude aktiviert; leer = 0."
    ws[f"{hinweis}5"] = ("AfA-Methode degressiv: par_AfADegressiv vom Restbuchwert, Wechsel zur "
                         "linearen AfA über die Restnutzungsdauer (1 / AfA-Satz), sobald höher.")
    ws.freeze_panes = "B2"


def _blatt_ruecklagen(wb) -> None:
    """§ 6b-Rücklagen: links je Verkauf, rechts der Spiegel je Jahr."""
    ws = wb.create_sheet("Rücklagen")
    erste = 2

    letzte = MAX_VERKAEUFE + 1
    _kopf(ws, 1, [s.ueberschrift for s in RUECKLAGE_SPALTEN])
    for i, s in enumerate(RUECKLAGE_SPALTEN, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'Rücklagen'!${bst}${erste}:${bst}${letzte}")
    for verkauf_nr, zeile in enumerate(range(erste, letzte + 1), start=1):
        rechnung = formeln.ruecklage_zeile(zeile, verkauf_nr)
        for i, s in enumerate(RUECKLAGE_SPALTEN, start=1):
            c = ws.cell(row=zeile, column=i, value=rechnung[s.key])
            c.number_format = s.format
            c.fill = FILL_BERECHNET

    versatz = len(RUECKLAGE_SPALTEN) + 1  # eine Spalte Abstand
    ws.column_dimensions[get_column_letter(versatz)].width = 3
    letzte = erste + prognosejahre() - 1
    for i, s in enumerate(RUECKLAGE_JAHR_SPALTEN, start=versatz + 1):
        bst = get_column_letter(i)
        c = ws.cell(row=1, column=i, value=s.ueberschrift)
        c.font, c.fill = FONT_KOPF, FILL_KOPF
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'Rücklagen'!${bst}${erste}:${bst}${letzte}")
    for zeile in range(erste, letzte + 1):
        rechnung = formeln.ruecklage_jahr_zeile(zeile, erstes_jahr=zeile == erste)
        for i, s in enumerate(RUECKLAGE_JAHR_SPALTEN, start=versatz + 1):
            c = ws.cell(row=zeile, column=i, value=rechnung[s.key])
            c.number_format = s.format
            c.fill = FILL_BERECHNET
    ws.row_dimensions[1].height = 45

    hinweis = letzte + 2
    ws.cell(row=hinweis, column=versatz + 1, value=(
        "Rücklage nur bei § 6b ja und Status OK im Blatt Verkäufe, gebildet zum Ende des "
        "Verkaufsjahrs aus den positiven Teilgewinnen Gebäude und G+B."))
    ws.cell(row=hinweis + 1, column=versatz + 1, value=(
        "Was bis zum Fristjahr nicht auf Neuobjekte übertragen ist, wird dort aufgelöst, "
        "plus Gewinnzuschlag je vollem Jahr."))
    ws.cell(row=hinweis + 2, column=versatz + 1, value=(
        "steuerwirksam = Veräußerungsgewinne − Einstellung + Auflösung + Zuschlag. "
        + HINWEIS_FINANZIERUNG))
    ws.freeze_panes = "B2"


TITEL_PLAN = SZ_A.kurz
TITEL_BASELINE = SZ_BASELINE.kurz


def _jahrestabellen(wb, ws, blatt: str, tabellen: list, hinweise: list) -> None:
    """Tabellen je Prognosejahr nebeneinander, je eine Spalte Abstand.

    tabellen: (Titel, Spalten, Formelfunktion(zeile, erstes_jahr)). Zeile 1 trägt
    den Titel, Zeile 2 die Kopfzeile, ab Zeile 3 je Jahr eine Zeile.
    """
    titel_zeile, kopf, erste = 1, 2, 3
    letzte = erste + prognosejahre() - 1
    versatz = 0
    for titel, spalten, formel in tabellen:
        ws.cell(row=titel_zeile, column=versatz + 1, value=titel).font = Font(bold=True)
        for i, s in enumerate(spalten, start=versatz + 1):
            bst = get_column_letter(i)
            c = ws.cell(row=kopf, column=i, value=s.ueberschrift)
            c.font, c.fill = FONT_KOPF, FILL_KOPF
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[bst].width = s.breite
            _name(wb, s.name, f"'{blatt}'!${bst}${erste}:${bst}${letzte}")
        for zeile in range(erste, letzte + 1):
            rechnung = formel(zeile, erstes_jahr=zeile == erste)
            for i, s in enumerate(spalten, start=versatz + 1):
                c = ws.cell(row=zeile, column=i, value=rechnung[s.key])
                c.number_format = s.format
                c.fill = FILL_BERECHNET
        versatz += len(spalten) + 1
        ws.column_dimensions[get_column_letter(versatz)].width = 3
    ws.row_dimensions[kopf].height = 45
    for i, text in enumerate(hinweise + [HINWEIS_FINANZIERUNG], start=letzte + 2):
        ws.cell(row=i, column=1, value=text)
    ws.freeze_panes = f"B{erste}"


def _blatt_liquiditaet(wb) -> None:
    """Steuer und Geldfluss je Jahr, je Szenario eine Tabelle."""
    ws = wb.create_sheet("Liquidität")
    _jahrestabellen(wb, ws, "Liquidität", [
        (sz.titel, liq_spalten(sz), formeln.liquiditaet_zeile(sz)) for sz in SZENARIEN
    ], [
        "Steuer = Bemessungsgrundlage × Grenzsteuersatz. Verluste werden vorgetragen und mit "
        "späteren Gewinnen verrechnet (ohne Mindestbesteuerung, ohne Rücktrag).",
        "steuerwirksam aus Verkauf und Rücklage: in A aus dem Rücklagenspiegel (sofort versteuerte "
        "Gewinne, Auflösung und Gewinnzuschlag), in B und C jeder Veräußerungsgewinn sofort.",
        "Zinsertrag = Liquidität am Vorjahresende × Rendite Alternativanlage; er ist steuerpflichtig. "
        "Negative Liquidität kostet denselben Satz.",
        "freier Mittelzufluss = Einnahmen − Ausgaben + Zins + Verkaufserlöse − Steuer − Kauf "
        "Neuobjekte.",
        "B: Neuobjekte mit Quelle-Rücklage entfallen, ihr Geld bleibt in der Alternativanlage. "
        "C: alle Neuobjekte werden gekauft, die AfA läuft von den vollen AK.",
        "Baseline: alle Bestandsobjekte werden über das ganze Raster gehalten, ohne Verkäufe und "
        "Neuobjekte.",
    ])


def _blatt_auswertung(wb) -> None:
    """Gesamt-GuV, Steuer, stille Reserven und Gesamtvermögen je Jahr und Szenario."""
    ws = wb.create_sheet("Auswertung")
    _jahrestabellen(wb, ws, "Auswertung", [
        (sz.titel, aus_spalten(sz), formeln.auswertung_zeile(sz)) for sz in SZENARIEN
    ], [
        "Gesamt-GuV vor Steuern = laufendes Ergebnis + steuerwirksam aus Verkauf und Rücklage "
        "+ Zinsertrag.",
        "stille Reserven = Verkehrswert − Buchwert (Gebäude + G+B) der Objekte im Bestand am "
        "Jahresende. In B und C ohne § 6b-Kürzung der Neuobjekte.",
        "latente Steuer = (stille Reserven + Rücklagenbestand − Verlustvortrag) × Grenzsteuersatz, "
        "mindestens 0: die Steuer, wenn alle Objekte zum Verkehrswert verkauft würden.",
        "Gesamtvermögen = Verkehrswert Bestand + Liquidität kumuliert.",
    ])


def _blatt_vergleich(wb) -> None:
    """Kennzahlen der Szenarien am Ende des Rasters und Endvermögen je Jahr."""
    ws = wb.create_sheet("Vergleich", 1)
    ws["A1"] = "Vergleich der Szenarien"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = ("A § 6b-Kette wie erfasst · B sofort versteuern, Kapital anlegen · C sofort "
                "versteuern, Neuobjekte trotzdem kaufen · Baseline alles halten.")
    ws["A3"] = ("Entscheidend ist das Endvermögen nach latenter Steuer. A − C zeigt die reine "
                "Wirkung von § 6b (Zins auf die gestundete Steuer), A − B die Frage Immobilie oder "
                "Geldanlage. " + HINWEIS_FINANZIERUNG)

    kopf = 5
    spalten = ([("Kennzahl", "vg_Kennzahl", 38)]
               + [(sz.kurz, f"vg_{sz.key}", 18) for sz in SZENARIEN]
               + [("A − B", "vg_DiffB", 16), ("A − C", "vg_DiffC", 16),
                  ("A − Baseline", "vg_DiffBaseline", 16), ("Erläuterung", None, 60)])
    _kopf(ws, kopf, [u for u, _, _ in spalten])
    ws.row_dimensions[kopf].height = 32
    erste = kopf + 1
    letzte = erste + len(VERGLEICH_KENNZAHLEN) - 1
    spalte_von = {sz.key: get_column_letter(2 + i) for i, sz in enumerate(SZENARIEN)}
    for zeile, (text, blatt, name, art, erlaeuterung) in enumerate(VERGLEICH_KENNZAHLEN,
                                                                   start=erste):
        werte = [text] + [formeln.vergleich_zeile(blatt, name, art, sz) for sz in SZENARIEN]
        a = spalte_von["A"]
        werte += [f"={a}{zeile}-{spalte_von[k]}{zeile}" for k in ("B", "C", "Baseline")]
        werte.append(erlaeuterung)
        for i, w in enumerate(werte, start=1):
            c = ws.cell(row=zeile, column=i, value=w)
            if 1 < i < len(werte):
                c.number_format = FMT_EURO
                c.fill = FILL_BERECHNET
        if zeile == erste:
            for i in range(1, len(werte)):
                ws.cell(row=zeile, column=i).font = Font(bold=True)
    for i, (_, name, breite) in enumerate(spalten, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = breite
        if name:
            _name(wb, name, f"'Vergleich'!${bst}${erste}:${bst}${letzte}")

    # Endvermögen je Jahr und Szenario, mit Diagramm
    titel = letzte + 2
    ws.cell(row=titel, column=1,
            value="Endvermögen nach latenter Steuer je Jahr").font = Font(bold=True)
    jkopf = titel + 1
    _kopf(ws, jkopf, ["Jahr"] + [sz.kurz for sz in SZENARIEN])
    ws.row_dimensions[jkopf].height = 32
    jerste, jletzte = jkopf + 1, jkopf + prognosejahre()
    for zeile in range(jerste, jletzte + 1):
        c = ws.cell(row=zeile, column=1,
                    value="=par_Startjahr" if zeile == jerste else f"=A{zeile - 1}+1")
        c.number_format, c.fill = FMT_JAHR, FILL_BERECHNET
        for i, sz in enumerate(SZENARIEN, start=2):
            c = ws.cell(row=zeile, column=i,
                        value=f"=SUMIFS({sz.aus}_VermoegenNetto,{sz.aus}_Jahr,A{zeile})")
            c.number_format, c.fill = FMT_EURO, FILL_BERECHNET
    _name(wb, "vgj_Jahr", f"'Vergleich'!$A${jerste}:$A${jletzte}")
    for i, sz in enumerate(SZENARIEN, start=2):
        bst = get_column_letter(i)
        _name(wb, f"vgj_{sz.key}", f"'Vergleich'!${bst}${jerste}:${bst}${jletzte}")

    chart = LineChart()
    chart.title = "Endvermögen nach latenter Steuer"
    chart.y_axis.title = "€"
    chart.y_axis.number_format = '#,##0'
    chart.x_axis.title = "Jahr"
    chart.add_data(Reference(ws, min_col=2, max_col=1 + len(SZENARIEN), min_row=jkopf,
                             max_row=jletzte), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=jerste, max_row=jletzte))
    farben = {"A": "305496", "B": "C00000", "C": "70AD47", "Baseline": "A5A5A5"}
    for serie, sz in zip(chart.series, SZENARIEN):
        serie.smooth = False
        serie.graphicalProperties.line.solidFill = farben[sz.key]
        serie.graphicalProperties.line.width = 28000
        if sz == SZ_BASELINE:
            serie.graphicalProperties.line.dashStyle = "dash"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"
    chart.height, chart.width = 12, 22
    ws.add_chart(chart, f"H{titel}")
    ws.freeze_panes = f"B{erste}"


def _blatt_uebersicht(wb) -> None:
    """Immobilienwert und Gesamtvermögen je Jahr: Plan mit Verkäufen gegen Halten ohne Verkauf."""
    ws = wb.create_sheet("Übersicht", 0)
    wb.active = 0
    ws["A1"] = "Übersicht: Immobilienwert und Gesamtvermögen"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = (HINWEIS_FINANZIERUNG + " Wert = Verkehrswert aktuell aus dem Objektblatt, "
                "fortgeschrieben mit der Wertsteigerung vom Parameterblatt.")
    ws["A3"] = ("Gesamtvermögen = Verkehrswert + kumulierte Liquidität nach Steuern − latente "
                "Steuer auf stille Reserven und Rücklage (Blatt Auswertung).")

    kennzahlen = [
        ("Objekte im Modell", "ueb_Objekte", '=SUMPRODUCT(--(obj_ID<>""))', FMT_ZAHL),
        ("davon ohne Verkehrswert (zählen mit 0)", "ueb_OhneWert",
         '=SUMPRODUCT((obj_ID<>"")*(obj_Verkehrswert=""))', FMT_ZAHL),
        ("geplante Verkäufe", "ueb_Verkaeufe", '=SUMPRODUCT(--(vk_ID<>""))', FMT_ZAHL),
        ("Neuobjekte im Modell", "ueb_Neuobjekte", "=SUM(ne_Gueltig)", FMT_ZAHL),
        ("Wertsteigerung p. a.", None, "=par_Wertsteig", FMT_PROZENT),
        ("Plausibilitätsprüfung (Blatt Prüfung)", "ueb_Pruefung", "=pr_Gesamt", FMT_TEXT),
    ]
    for zeile, (text, name, formel, fmt) in enumerate(kennzahlen, start=4):
        ws.cell(row=zeile, column=1, value=text)  # läuft über die leeren Spalten B und C
        c = ws.cell(row=zeile, column=4, value=formel)
        c.fill = FILL_BERECHNET
        c.number_format = fmt
        if name:
            _name(wb, name, f"'Übersicht'!$D${zeile}")
    # Hinweis, solange Verkehrswerte fehlen
    ws.conditional_formatting.add(
        "D5", FormulaRule(formula=["D5>0"], fill=FILL_FEHLER))
    ws.conditional_formatting.add(
        "D9", FormulaRule(formula=['D9<>"OK"'], fill=FILL_FEHLER))

    kopf = 11
    ws.cell(row=kopf - 1, column=2, value="Immobilien (Verkehrswert)").font = Font(bold=True)
    ws.cell(row=kopf - 1, column=5,
            value="Gesamtvermögen nach latenter Steuer").font = Font(bold=True)
    _kopf(ws, kopf, ["Jahr", TITEL_BASELINE, TITEL_PLAN, "Differenz Plan − Baseline",
                     TITEL_BASELINE, TITEL_PLAN, "Differenz Plan − Baseline"])
    ws.row_dimensions[kopf].height = 32
    erste = kopf + 1
    letzte = erste + prognosejahre()  # Basisjahr plus Prognosejahre
    for zeile in range(erste, letzte + 1):
        if zeile == erste:
            # Ende Basisjahr: Stand laut Objektblatt, noch keine Liquidität
            buchwert = 'SUMIFS(obj_Restbuchwert,obj_ID,"<>")+SUMIFS(obj_AKGuB,obj_ID,"<>")'
            werte = ["=par_Basisjahr", "=SUM(obj_Verkehrswert)", f"=B{zeile}",
                     f"=C{zeile}-B{zeile}",
                     f"=B{zeile}-MAX(B{zeile}-({buchwert}),0)*par_Steuersatz", f"=E{zeile}"]
        else:
            werte = [f"=A{zeile - 1}+1",
                     f"=SUMIFS(prg_Verkehrswert,prg_Jahr,A{zeile},prg_Neu,0)",
                     f"=SUMIFS(prg_Verkehrswert,prg_Jahr,A{zeile},prg_Bestand,1)",
                     f"=C{zeile}-B{zeile}",
                     f"=SUMIFS(asb_VermoegenNetto,asb_Jahr,A{zeile})",
                     f"=SUMIFS(aus_VermoegenNetto,aus_Jahr,A{zeile})"]
        werte.append(f"=F{zeile}-E{zeile}")
        for spalte, formel in enumerate(werte, start=1):
            c = ws.cell(row=zeile, column=spalte, value=formel)
            c.number_format = FMT_JAHR if spalte == 1 else FMT_EURO
            c.fill = FILL_BERECHNET
    for name, bst in (("ueb_Jahr", "A"), ("ueb_Baseline", "B"), ("ueb_Plan", "C"),
                      ("ueb_Differenz", "D"), ("ueb_VermBaseline", "E"), ("ueb_VermPlan", "F"),
                      ("ueb_VermDifferenz", "G")):
        _name(wb, name, f"'Übersicht'!${bst}${erste}:${bst}${letzte}")

    for spalte, breite in zip("ABCDEFG", (10, 20, 20, 20, 20, 20, 20)):
        ws.column_dimensions[spalte].width = breite
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    for i, (titel, spalte_baseline) in enumerate(
            (("Verkehrswert am Jahresende", 2),
             ("Gesamtvermögen nach latenter Steuer", 5))):
        ws.add_chart(_diagramm(ws, titel, spalte_baseline, kopf, erste, letzte), f"I{4 + i * 26}")


def _diagramm(ws, titel: str, spalte_baseline: int, kopf: int, erste: int, letzte: int):
    """Liniendiagramm Baseline (gestrichelt) gegen Plan aus zwei benachbarten Spalten."""
    chart = LineChart()
    chart.title = titel
    chart.y_axis.title = "€"
    chart.y_axis.number_format = '#,##0'
    chart.y_axis.majorGridlines = None
    chart.x_axis.title = "Jahr"
    chart.add_data(Reference(ws, min_col=spalte_baseline, max_col=spalte_baseline + 1,
                             min_row=kopf, max_row=letzte), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=erste, max_row=letzte))
    baseline, plan = chart.series
    baseline.graphicalProperties.line.solidFill = "A5A5A5"
    baseline.graphicalProperties.line.dashStyle = "dash"
    plan.graphicalProperties.line.solidFill = "305496"
    for serie in chart.series:
        serie.smooth = False
        serie.graphicalProperties.line.width = 28000
    # openpyxl blendet die Achsen sonst in neueren Excel-Versionen aus
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"
    chart.height, chart.width = 12, 22
    return chart


def _blatt_pruefung(wb) -> None:
    """Plausibilitätsprüfungen als Formeln: je Prüfung Anzahl betroffener Zeilen und Ergebnis.

    Das Makro modPruefung liest nur das Ergebnis; geprüft wird im Blatt, damit Fehler
    auch ohne Makros sichtbar bleiben.
    """
    ws = wb.create_sheet("Prüfung")
    ws["A1"] = "Plausibilitätsprüfung"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = HINWEIS_FINANZIERUNG
    summen = formeln.pruefung_summen()
    for zeile, (name, text) in enumerate([("pr_Gesamt", "Gesamtergebnis"),
                                          ("pr_Fehler", "Fehler"),
                                          ("pr_Warnungen", "Warnungen")], start=3):
        ws.cell(row=zeile, column=1, value=text).font = Font(bold=True)
        c = ws.cell(row=zeile, column=2, value=summen[name])
        c.fill = FILL_BERECHNET
        c.number_format = FMT_TEXT if name == "pr_Gesamt" else FMT_ZAHL
        _name(wb, name, f"'Prüfung'!$B${zeile}")
    ws.conditional_formatting.add("B3", FormulaRule(formula=['B3<>"OK"'], fill=FILL_FEHLER))

    _kopf(ws, 7, ["Nr", "Prüfung", "Art", "betroffen", "Ergebnis", "wo nachsehen"])
    erste = 8
    letzte = erste + len(PRUEFUNGEN) - 1
    anzahl = formeln.pruefung_anzahl()
    for nr, (zeile, p) in enumerate(zip(range(erste, letzte + 1), PRUEFUNGEN), start=1):
        ws.cell(row=zeile, column=1, value=nr)
        ws.cell(row=zeile, column=2, value=p.bezeichnung).alignment = Alignment(wrap_text=True)
        ws.cell(row=zeile, column=3, value=p.art)
        for spalte, formel in ((4, anzahl[p.key]), (5, formeln.pruefung_ergebnis(zeile))):
            c = ws.cell(row=zeile, column=spalte, value=formel)
            c.fill = FILL_BERECHNET
        ws.cell(row=zeile, column=6, value=p.wo).alignment = Alignment(wrap_text=True)
    for name, bst in (("pr_Bezeichnung", "B"), ("pr_Art", "C"), ("pr_Anzahl", "D"),
                      ("pr_Ergebnis", "E")):
        _name(wb, name, f"'Prüfung'!${bst}${erste}:${bst}${letzte}")
    for art, fill in ((FEHLER, FILL_FEHLER), (WARNUNG, FILL_WARNUNG)):
        ws.conditional_formatting.add(
            f"E{erste}:E{letzte}", FormulaRule(formula=[f'E{erste}="{art}"'], fill=fill))

    hinweise = [
        f"{FEHLER}: Eingabe unvollständig oder unzulässig; die Zeile rechnet nicht oder nur "
        f"teilweise mit. {WARNUNG}: rechnet, ist aber steuerlich ungünstig oder fachlich zu "
        f"prüfen. {HINWEIS}: zur Kenntnis, zählt nicht ins Gesamtergebnis.",
        "Das Gesamtergebnis steht auch auf dem Parameterblatt und in der Übersicht. Das Makro "
        "„Plausibilität prüfen“ rechnet neu und listet die auffälligen Prüfungen auf.",
        "Je Objekt, Verkauf und Neuobjekt steht der Grund in der Statusspalte des Blatts. "
        "AK Gebäude und AK G+B werden getrennt erfasst, ihre Summe ist der Kaufpreis; beim "
        "Neuobjekt ergibt der Anteil G+B mit dem Rest Gebäude immer den Kaufpreis.",
    ]
    for i, text in enumerate(hinweise, start=letzte + 2):
        ws.cell(row=i, column=1, value=text)
    for spalte, breite in zip("ABCDEF", (16, 70, 10, 10, 10, 50)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A8"


def _blatt_varianten(wb) -> None:
    """Festgehaltene Ergebnisse: das Makro „Variante festhalten“ schreibt je Lauf eine Zeile."""
    ws = wb.create_sheet("Varianten")
    ws["A1"] = "Varianten: festgehaltene Ergebnisse"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = (HINWEIS_FINANZIERUNG + " Feste Werte, keine Formeln: Eingaben ändern (etwa "
                "Verkaufsjahr oder Preis), dann erneut festhalten und die Zeilen vergleichen.")
    kopf = 4
    _kopf(ws, kopf, VARIANTEN_KOPF)
    ws.row_dimensions[kopf].height = 32
    erste, letzte = kopf + 1, kopf + MAX_VARIANTEN
    _name(wb, "var_Tabelle",
          f"Varianten!$A${erste}:${get_column_letter(len(VARIANTEN_KOPF))}${letzte}")
    for name, bst in zip(("var_Bezeichnung", "var_Datum", "var_A", "var_B", "var_C",
                          "var_Baseline", "var_DiffB", "var_DiffC", "var_Steuer",
                          "var_Pruefung"), "ABCDEFGHIJ"):
        _name(wb, name, f"Varianten!${bst}${erste}:${bst}${letzte}")
    for zeile in range(erste, letzte + 1):
        ws.cell(row=zeile, column=2).number_format = "DD.MM.YYYY HH:MM"
        for spalte in range(3, len(VARIANTEN_KOPF)):
            ws.cell(row=zeile, column=spalte).number_format = FMT_EURO
    for spalte, breite in zip("ABCDEFGHIJ", (28, 16, 16, 16, 16, 18, 14, 14, 16, 16)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = f"B{erste}"


def erstelle_mappe(modell: Modell) -> Workbook:
    wb = Workbook()
    _blatt_parameter(wb, modell)
    _blatt_objekte(wb, modell)
    _blatt_verkaeufe(wb, modell)
    _blatt_neuobjekte(wb, modell)
    _blatt_prognose(wb)
    _blatt_ruecklagen(wb)
    _blatt_liquiditaet(wb)
    _blatt_auswertung(wb)
    _blatt_uebersicht(wb)
    _blatt_vergleich(wb)
    _blatt_pruefung(wb)
    _blatt_varianten(wb)
    wb.code_name = CODENAME_MAPPE
    for ws in wb.worksheets:
        ws.sheet_properties.codeName = CODENAMEN[ws.title]
    return wb
