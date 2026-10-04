"""Schreibt die Arbeitsmappe: Blätter, Kopfzeilen, benannte Bereiche, Validierung."""

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import formeln
from .modelle import (MAX_OBJEKTE, OBJEKT_FELDER, PARAMETER, PROGNOSE_SPALTEN,
                      STATUS_NAME, STATUS_UEBERSCHRIFT, Modell)

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
            dv.add(f"{bst}{erste}:{bst}{letzte}")

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

    # Status rot, sobald er gesetzt ist und nicht OK lautet
    ws.conditional_formatting.add(
        f"{st}{erste}:{st}{letzte}",
        FormulaRule(formula=[f'AND({st}{erste}<>"",{st}{erste}<>"OK")'], fill=FILL_FEHLER),
    )

    for zeile, obj in enumerate(modell.objekte, start=erste):
        for i, f in enumerate(OBJEKT_FELDER, start=1):
            wert = getattr(obj, f.key)
            if wert is not None:
                ws.cell(row=zeile, column=i, value=wert)

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
    _blatt_prognose(wb, modell)
    return wb
