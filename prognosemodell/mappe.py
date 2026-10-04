"""Schreibt die Arbeitsmappe: Blätter, Kopfzeilen, benannte Bereiche, Validierung."""

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import formeln
from .modelle import (MAX_OBJEKTE, MAX_VERKAEUFE, OBJEKT_FELDER, PARAMETER, PROGNOSE_SPALTEN,
                      STATUS_NAME, STATUS_UEBERSCHRIFT, VERKAUF_FELDER, VERKAUF_SPALTEN,
                      VERKAUF_STATUS_NAME, Modell)

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


def _blatt_parameter(wb, modell: Modell) -> None:
    ws = wb.active
    ws.title = "Parameter"
    ws["A1"] = "Parameter"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = HINWEIS_FINANZIERUNG
    _kopf(ws, 4, ["Bezeichnung", "Wert", "Name", "Erläuterung"])

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
        if p.minimum is not None or p.maximum is not None:
            dv = DataValidation(type="decimal", operator="between",
                                formula1=str(p.minimum), formula2=str(p.maximum),
                                showErrorMessage=True, errorTitle="Ungültiger Wert",
                                error=f"{p.bezeichnung}: erlaubt sind {p.minimum:.0%} "
                                      f"bis {p.maximum:.0%}.")
            ws.add_data_validation(dv)
            dv.add(c.coordinate)
        if p.name.startswith("par_Status"):
            ws.conditional_formatting.add(
                c.coordinate,
                FormulaRule(formula=[f'{c.coordinate}<>"OK"'], fill=FILL_FEHLER),
            )

    for spalte, breite in zip("ABCD", (36, 14, 22, 60)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A5"


def _eingabespalten(wb, ws, blatt: str, felder, erste: int, letzte: int) -> None:
    """Eingabespalten: Breite, benannter Bereich, Format, gelbe Füllung, Validierung."""
    for i, f in enumerate(felder, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = f.breite
        _name(wb, f.name, f"'{blatt}'!${bst}${erste}:${bst}${letzte}")
        bereich = f"{bst}{erste}:{bst}{letzte}"
        if f.minimum is not None or f.maximum is not None:
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
        if f.auswahl:
            dv = DataValidation(type="list", formula1='"' + ",".join(f.auswahl) + '"',
                                allow_blank=True, showErrorMessage=True)
            ws.add_data_validation(dv)
            dv.add(bereich)
        for zeile in range(erste, letzte + 1):
            c = ws.cell(row=zeile, column=i)
            c.number_format = f.format
            c.fill = FILL_EINGABE


def _datensaetze(ws, felder, datensaetze, erste: int) -> None:
    for zeile, satz in enumerate(datensaetze, start=erste):
        for i, f in enumerate(felder, start=1):
            wert = getattr(satz, f.key)
            if wert is not None:
                ws.cell(row=zeile, column=i, value=wert)


def _status_rot(ws, st: str, erste: int, letzte: int) -> None:
    """Status rot, sobald er gesetzt ist und nicht OK lautet."""
    ws.conditional_formatting.add(
        f"{st}{erste}:{st}{letzte}",
        FormulaRule(formula=[f'AND({st}{erste}<>"",{st}{erste}<>"OK")'], fill=FILL_FEHLER),
    )


def _blatt_objekte(wb, modell: Modell) -> None:
    ws = wb.create_sheet("Objekte")
    erste, letzte = 2, MAX_OBJEKTE + 1
    status_spalte = len(OBJEKT_FELDER) + 1
    _kopf(ws, 1, [f.ueberschrift for f in OBJEKT_FELDER] + [STATUS_UEBERSCHRIFT])
    ws.row_dimensions[1].height = 32
    _eingabespalten(wb, ws, "Objekte", OBJEKT_FELDER, erste, letzte)

    st = get_column_letter(status_spalte)
    ws.column_dimensions[st].width = 26
    _name(wb, STATUS_NAME, f"Objekte!${st}${erste}:${st}${letzte}")
    _name(wb, "obj_Basis", f"Objekte!$A${erste}:${st}${letzte}")
    for zeile in range(erste, letzte + 1):
        c = ws.cell(row=zeile, column=status_spalte, value=formeln.status_objekt(zeile))
        c.fill = FILL_BERECHNET
    _status_rot(ws, st, erste, letzte)

    _datensaetze(ws, OBJEKT_FELDER, modell.objekte, erste)
    ws.freeze_panes = "B2"


def _blatt_verkaeufe(wb, modell: Modell) -> None:
    """Je Zeile ein geplanter Verkauf; Aufteilung des Erlöses in Buchwert und Gewinn."""
    blatt = "Verkäufe"
    ws = wb.create_sheet(blatt)
    erste, letzte = 2, MAX_VERKAEUFE + 1
    sp = formeln.verkauf_spalten()
    _kopf(ws, 1, [f.ueberschrift for f in VERKAUF_FELDER] + [STATUS_UEBERSCHRIFT]
          + [s.ueberschrift for s in VERKAUF_SPALTEN])
    ws.row_dimensions[1].height = 45
    _eingabespalten(wb, ws, blatt, VERKAUF_FELDER, erste, letzte)

    st = sp["status"]
    ws.column_dimensions[st].width = 28
    _name(wb, VERKAUF_STATUS_NAME, f"'{blatt}'!${st}${erste}:${st}${letzte}")
    for s in VERKAUF_SPALTEN:
        bst = sp[s.key]
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'{blatt}'!${bst}${erste}:${bst}${letzte}")

    for zeile in range(erste, letzte + 1):
        zelle_formeln = formeln.verkauf_zeile(zeile)
        c = ws.cell(row=zeile, column=len(VERKAUF_FELDER) + 1, value=zelle_formeln["status"])
        c.fill = FILL_BERECHNET
        for i, s in enumerate(VERKAUF_SPALTEN, start=len(VERKAUF_FELDER) + 2):
            c = ws.cell(row=zeile, column=i, value=zelle_formeln[s.key])
            c.number_format = s.format
            c.fill = FILL_BERECHNET
    _status_rot(ws, st, erste, letzte)

    _datensaetze(ws, VERKAUF_FELDER, modell.verkaeufe, erste)
    ws.cell(row=1, column=len(VERKAUF_FELDER) + len(VERKAUF_SPALTEN) + 3,
            value="Verkauf zum Jahresende; das Objekt rechnet ab dem Folgejahr nicht mehr. "
                  + HINWEIS_FINANZIERUNG)
    ws.freeze_panes = "B2"


def prognosejahre(modell: Modell) -> int:
    standard = next(p.wert for p in PARAMETER if p.name == "par_Prognosejahre")
    return int(modell.parameter.get("par_Prognosejahre", standard))


def _blatt_prognose(wb, modell: Modell) -> None:
    """Long-Format: je Objektzeile ein Block mit einer Zeile je Prognosejahr."""
    ws = wb.create_sheet("Prognose")
    jahre = prognosejahre(modell)
    erste, letzte = 2, MAX_OBJEKTE * jahre + 1
    _kopf(ws, 1, [s.ueberschrift for s in PROGNOSE_SPALTEN])
    ws.row_dimensions[1].height = 32

    for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"Prognose!${bst}${erste}:${bst}${letzte}")

    zeile = erste
    for objekt_zeile in range(2, MAX_OBJEKTE + 2):
        for jahr_index in range(jahre):
            zelle_formeln = formeln.prognose_zeile(zeile, objekt_zeile, jahr_index)
            for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
                c = ws.cell(row=zeile, column=i, value=zelle_formeln[s.key])
                c.number_format = s.format
            zeile += 1

    ws.cell(row=1, column=len(PROGNOSE_SPALTEN) + 2, value=HINWEIS_FINANZIERUNG)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(PROGNOSE_SPALTEN))}{letzte}"
    ws.freeze_panes = "C2"
    # Rechenblatt: schützen, Filtern bleibt erlaubt (ohne Kennwort)
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False


def erstelle_mappe(modell: Modell) -> Workbook:
    wb = Workbook()
    _blatt_parameter(wb, modell)
    _blatt_objekte(wb, modell)
    _blatt_verkaeufe(wb, modell)
    _blatt_prognose(wb, modell)
    return wb
