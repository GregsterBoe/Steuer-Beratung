"""Schreibt die Arbeitsmappe: Blätter, Kopfzeilen, benannte Bereiche, Validierung."""

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import formeln
from .modelle import (FMT_EURO, FMT_JAHR, FMT_PROZENT, FMT_ZAHL, MAX_OBJEKTE, MAX_VERKAEUFE, OBJEKT_FELDER,
                      PARAMETER, PROGNOSE_SPALTEN, STATUS_NAME, STATUS_UEBERSCHRIFT,
                      VERKAUF_FELDER, VERKAUF_SPALTEN, VERKAUF_STATUS_NAME, Modell,
                      prognosejahre)

HINWEIS_FINANZIERUNG = "Alle Werte vor Finanzierung (ohne Zins und Tilgung)."

FONT_TITEL = Font(bold=True, size=14)
FONT_KOPF = Font(bold=True, color="FFFFFF")
FILL_KOPF = PatternFill("solid", fgColor="305496")
FILL_EINGABE = PatternFill("solid", fgColor="FFF2CC")   # gelb = hier wird getippt
FILL_BERECHNET = PatternFill("solid", fgColor="E7E6E6")  # grau = Formel
FILL_FEHLER = PatternFill("solid", fgColor="F8CBAD")


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


def _blatt_parameter(wb) -> None:
    ws = wb.active
    ws.title = "Parameter"
    ws["A1"] = "Parameter"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = HINWEIS_FINANZIERUNG
    _kopf(ws, 4, ["Bezeichnung", "Wert", "Name", "Erläuterung"])

    for zeile, p in enumerate(PARAMETER, start=5):
        ws.cell(row=zeile, column=1, value=p.bezeichnung)
        c = ws.cell(row=zeile, column=2, value=p.wert)
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
    """Je Zeile des Objektblatts ein Block mit einer Zeile je Prognosejahr."""
    ws = wb.create_sheet("Prognose")
    jahre = prognosejahre()
    erste, letzte = 2, MAX_OBJEKTE * jahre + 1
    _kopf(ws, 1, [s.ueberschrift for s in PROGNOSE_SPALTEN])
    ws.row_dimensions[1].height = 32

    for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"Prognose!${bst}${erste}:${bst}${letzte}")

    for objekt_nr in range(1, MAX_OBJEKTE + 1):
        for j in range(jahre):
            zeile = erste + (objekt_nr - 1) * jahre + j
            formeln_zeile = formeln.prognose_zeile(zeile, objekt_nr, erstes_jahr=j == 0)
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
    ws.freeze_panes = "B2"


def _blatt_uebersicht(wb) -> None:
    """Gesamtwert aller Objekte je Jahr: Plan mit Verkäufen gegen Halten ohne Verkauf."""
    ws = wb.create_sheet("Übersicht", 0)
    wb.active = 0
    ws["A1"] = "Übersicht: Gesamtwert der Immobilien"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = (HINWEIS_FINANZIERUNG + " Wert = Verkehrswert aktuell aus dem Objektblatt, "
                "fortgeschrieben mit der Wertsteigerung vom Parameterblatt.")

    kennzahlen = [
        ("Objekte im Modell", "ueb_Objekte", '=SUMPRODUCT(--(obj_ID<>""))', FMT_ZAHL),
        ("davon ohne Verkehrswert (zählen mit 0)", "ueb_OhneWert",
         '=SUMPRODUCT((obj_ID<>"")*(obj_Verkehrswert=""))', FMT_ZAHL),
        ("geplante Verkäufe", "ueb_Verkaeufe", '=SUMPRODUCT(--(vk_ID<>""))', FMT_ZAHL),
        ("Wertsteigerung p. a.", None, "=par_Wertsteig", FMT_PROZENT),
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

    kopf = 9
    _kopf(ws, kopf, ["Jahr", "Baseline: alles halten", "Plan: mit Verkäufen",
                     "Differenz Plan − Baseline"])
    ws.row_dimensions[kopf].height = 32
    erste = kopf + 1
    letzte = erste + prognosejahre()  # Basisjahr plus Prognosejahre
    for zeile in range(erste, letzte + 1):
        if zeile == erste:  # Ende Basisjahr: Stand laut Objektblatt
            werte = ["=par_Basisjahr", "=SUM(obj_Verkehrswert)", f"=B{zeile}"]
        else:
            werte = [f"=A{zeile - 1}+1",
                     f"=SUMIFS(prg_Verkehrswert,prg_Jahr,A{zeile})",
                     f"=SUMIFS(prg_Verkehrswert,prg_Jahr,A{zeile},prg_Bestand,1)"]
        werte.append(f"=C{zeile}-B{zeile}")
        for spalte, (formel, fmt) in enumerate(
                zip(werte, (FMT_JAHR, FMT_EURO, FMT_EURO, FMT_EURO)), start=1):
            c = ws.cell(row=zeile, column=spalte, value=formel)
            c.number_format = fmt
            c.fill = FILL_BERECHNET
    for name, bst in (("ueb_Jahr", "A"), ("ueb_Baseline", "B"), ("ueb_Plan", "C"),
                      ("ueb_Differenz", "D")):
        _name(wb, name, f"'Übersicht'!${bst}${erste}:${bst}${letzte}")

    for spalte, breite in zip("ABCD", (10, 22, 22, 24)):
        ws.column_dimensions[spalte].width = breite
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    chart = LineChart()
    chart.title = "Verkehrswert am Jahresende"
    chart.y_axis.title = "€"
    chart.y_axis.number_format = '#,##0'
    chart.y_axis.majorGridlines = None
    chart.x_axis.title = "Jahr"
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=kopf, max_row=letzte),
                   titles_from_data=True)
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
    ws.add_chart(chart, "F4")


def erstelle_mappe(modell: Modell) -> Workbook:
    wb = Workbook()
    _blatt_parameter(wb)
    _blatt_objekte(wb, modell)
    _blatt_verkaeufe(wb, modell)
    _blatt_prognose(wb)
    _blatt_uebersicht(wb)
    return wb
