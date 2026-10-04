"""Schreibt die Arbeitsmappe: Blätter, Kopfzeilen, benannte Bereiche, Validierung."""

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import formeln
from .modelle import (AUSWERTUNG_SPALTEN, CODENAME_MAPPE, CODENAMEN, EINGABE_NAME, FEHLER,
                      FMT_EURO, FMT_TEXT, FMT_ZAHL, HINWEIS, LIQUIDITAET_SPALTEN,
                      MAX_NEUOBJEKTE, MAX_OBJEKTE, MAX_VERKAEUFE, NEU_FELDER, NEU_SPALTEN,
                      NEU_STATUS_NAME, OBJEKT_FELDER, PARAMETER, PROGNOSE_SPALTEN, PRUEFUNGEN,
                      RUECKLAGE_6B_GEBILDET, RUECKLAGE_JAHR_SPALTEN, RUECKLAGE_SPALTEN,
                      STATUS_NAME, STATUS_UEBERSCHRIFT, SZENARIEN, SZENARIO_A,
                      VERGLEICH_KENNZAHLEN, VERKAUF_FELDER, VERKAUF_SPALTEN,
                      VERKAUF_STATUS_NAME, WARNUNG, Modell)

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
    ws["F3"] = ("Steuerung per Makro (Datei als .xlsm, Makros aktivieren): Die Schaltflächen "
                "darunter legt die Mappe beim Öffnen an; ohne Schaltflächen Alt+F8.")
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
        if f.auswahl or f.auswahl_bereich:
            liste = f.auswahl_bereich or '"' + ",".join(f.auswahl) + '"'
            dv = DataValidation(type="list", formula1=liste, allow_blank=True,
                                showErrorMessage=True)
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


def _status_rot(ws, st: str, erste: int, letzte: int, ausnahme: str = "") -> None:
    """Status rot, sobald er gesetzt ist und nicht OK lautet (und nicht die Ausnahme ist)."""
    bedingungen = [f'{st}{erste}<>""', f'{st}{erste}<>"OK"']
    if ausnahme:
        bedingungen.append(f'{st}{erste}<>"{ausnahme}"')
    ws.conditional_formatting.add(
        f"{st}{erste}:{st}{letzte}",
        FormulaRule(formula=[f"AND({','.join(bedingungen)})"], fill=FILL_FEHLER),
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
    letzte_eingabe = get_column_letter(len(OBJEKT_FELDER))
    _name(wb, EINGABE_NAME, f"Objekte!$A${erste}:${letzte_eingabe}${letzte}")
    for zeile in range(erste, letzte + 1):
        c = ws.cell(row=zeile, column=status_spalte, value=formeln.status_objekt(zeile))
        c.fill = FILL_BERECHNET
    _status_rot(ws, st, erste, letzte)

    _datensaetze(ws, OBJEKT_FELDER, modell.objekte, erste)
    ws.freeze_panes = "B2"


def _eingabeblatt(wb, blatt: str, felder, status_name: str, spalten, buchstaben: dict,
                  letzte: int, zeilen_formeln, datensaetze, hinweis: str,
                  status_ausnahme: str = "") -> None:
    """Eingabeblatt mit Statusspalte und anschließenden Formelspalten (Verkäufe, Neuobjekte)."""
    ws = wb.create_sheet(blatt)
    erste = 2
    _kopf(ws, 1, [f.ueberschrift for f in felder] + [STATUS_UEBERSCHRIFT]
          + [s.ueberschrift for s in spalten])
    ws.row_dimensions[1].height = 45
    _eingabespalten(wb, ws, blatt, felder, erste, letzte)

    st = buchstaben["status"]
    ws.column_dimensions[st].width = 28
    _name(wb, status_name, f"'{blatt}'!${st}${erste}:${st}${letzte}")
    for s in spalten:
        bst = buchstaben[s.key]
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'{blatt}'!${bst}${erste}:${bst}${letzte}")

    for zeile in range(erste, letzte + 1):
        zelle_formeln = zeilen_formeln(zeile)
        c = ws[f"{st}{zeile}"]
        c.value = zelle_formeln["status"]
        c.fill = FILL_BERECHNET
        for s in spalten:
            c = ws[f"{buchstaben[s.key]}{zeile}"]
            c.value = zelle_formeln[s.key]
            c.number_format = s.format
            c.fill = FILL_BERECHNET
    _status_rot(ws, st, erste, letzte, status_ausnahme)

    _datensaetze(ws, felder, datensaetze, erste)
    ws.cell(row=1, column=len(felder) + len(spalten) + 3, value=hinweis)
    ws.freeze_panes = "B2"


def _blatt_verkaeufe(wb, modell: Modell) -> None:
    """Je Zeile ein geplanter Verkauf; Aufteilung des Erlöses in Buchwert und Gewinn."""
    _eingabeblatt(wb, "Verkäufe", VERKAUF_FELDER, VERKAUF_STATUS_NAME, VERKAUF_SPALTEN,
                  formeln.verkauf_spalten(), MAX_VERKAEUFE + 1, formeln.verkauf_zeile,
                  modell.verkaeufe,
                  "Verkauf zum Jahresende; das Objekt rechnet ab dem Folgejahr nicht mehr. "
                  + HINWEIS_FINANZIERUNG)


def _blatt_neuobjekte(wb, modell: Modell) -> None:
    """Je Zeile ein Reinvestitionsobjekt; Übertrag der Rücklage mindert die AfA-Basis."""
    _eingabeblatt(wb, "Neuobjekte", NEU_FELDER, NEU_STATUS_NAME, NEU_SPALTEN,
                  formeln.neu_spalten(), MAX_NEUOBJEKTE + 1, formeln.neu_zeile,
                  modell.neuobjekte,
                  "Kauf zum Jahresende; Miete, Erhaltung und AfA ab dem Folgejahr. Gebäude-"
                  "Rücklage nur auf Gebäude, G+B-Rücklage zuerst auf G+B, Rest auf Gebäude; "
                  "übertragen wird so viel wie möglich, obere Zeilen zuerst. In Szenario B "
                  "entfallen Neuobjekte mit Quelle-Rücklage, in Szenario C werden sie ohne "
                  "Übertrag gekauft. " + HINWEIS_FINANZIERUNG,
                  formeln.NEU_SZENARIO_B)


def prognosejahre(modell: Modell) -> int:
    standard = next(p.wert for p in PARAMETER if p.name == "par_Prognosejahre")
    return int(modell.parameter.get("par_Prognosejahre", standard))


def _blatt_prognose(wb, modell: Modell) -> None:
    """Long-Format: je Objektzeile ein Block mit einer Zeile je Prognosejahr.

    Erst alle Bestandsobjekte, darunter die Neuobjekte. prg_* läuft über beide Teile,
    prgb_* nur über den Bestand (für das Blatt Verkäufe, ohne Zirkelbezug).
    """
    ws = wb.create_sheet("Prognose")
    jahre = prognosejahre(modell)
    erste = 2
    letzte_bestand = MAX_OBJEKTE * jahre + 1
    letzte = letzte_bestand + MAX_NEUOBJEKTE * jahre
    _kopf(ws, 1, [s.ueberschrift for s in PROGNOSE_SPALTEN])
    ws.row_dimensions[1].height = 32

    for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
        bst = get_column_letter(i)
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"Prognose!${bst}${erste}:${bst}${letzte}")
        _name(wb, s.name.replace("prg_", "prgb_"),
              f"Prognose!${bst}${erste}:${bst}${letzte_bestand}")

    bloecke = [(formeln.prognose_zeile, z) for z in range(2, MAX_OBJEKTE + 2)]
    bloecke += [(formeln.prognose_neu_zeile, z) for z in range(2, MAX_NEUOBJEKTE + 2)]
    zeile = erste
    for zeilen_formeln, eingabe_zeile in bloecke:
        for jahr_index in range(jahre):
            zelle_formeln = zeilen_formeln(zeile, eingabe_zeile, jahr_index)
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


def _rechenspalten(wb, ws, blatt: str, spalten, buchstaben: dict, prefix: str,
                   erste: int, letzte: int, zeilen_formeln) -> None:
    """Formelspalten mit benanntem Bereich; zeilen_formeln(zeile, index) liefert die Formeln."""
    for s in spalten:
        bst = buchstaben[prefix + s.key]
        ws.column_dimensions[bst].width = s.breite
        _name(wb, s.name, f"'{blatt}'!${bst}${erste}:${bst}${letzte}")
    for index, zeile in enumerate(range(erste, letzte + 1)):
        zelle_formeln = zeilen_formeln(zeile, index)
        for s in spalten:
            c = ws[f"{buchstaben[prefix + s.key]}{zeile}"]
            c.value = zelle_formeln[s.key]
            c.number_format = s.format
            c.fill = FILL_BERECHNET


def _blatt_ruecklagen(wb, modell: Modell) -> None:
    """Links je Verkauf die § 6b-Prüfung und Rücklage, rechts der Spiegel je Jahr."""
    blatt = "Rücklagen"
    ws = wb.create_sheet(blatt)
    erste = 2
    sp = formeln.ruecklage_spalten()
    _kopf(ws, 1, [s.ueberschrift for s in RUECKLAGE_SPALTEN] + [None]
          + [s.ueberschrift for s in RUECKLAGE_JAHR_SPALTEN])
    leer = ws.cell(row=1, column=len(RUECKLAGE_SPALTEN) + 1)
    leer.fill = PatternFill(fill_type=None)
    ws.column_dimensions[leer.column_letter].width = 3
    ws.row_dimensions[1].height = 45

    _rechenspalten(wb, ws, blatt, RUECKLAGE_SPALTEN, sp, "", erste, MAX_VERKAEUFE + 1,
                   lambda zeile, _i: formeln.ruecklage_zeile(zeile))
    _rechenspalten(wb, ws, blatt, RUECKLAGE_JAHR_SPALTEN, sp, "j_", erste,
                   prognosejahre(modell) + 1, formeln.ruecklage_jahr_zeile)

    st = sp["status"]
    ws.conditional_formatting.add(
        f"{st}{erste}:{st}{MAX_VERKAEUFE + 1}",
        FormulaRule(formula=[f'{st}{erste}="Vorbesitzzeit zu kurz"'], fill=FILL_FEHLER),
    )
    hinweis = len(RUECKLAGE_SPALTEN) + len(RUECKLAGE_JAHR_SPALTEN) + 3
    ws.cell(row=1, column=hinweis,
            value=f"Zeile n gehört zu Zeile n im Blatt Verkäufe. Rücklage nur bei 6b-Nutzung ja, "
                  f"Status „{RUECKLAGE_6B_GEBILDET}“; je Wirtschaftsgut nur Gewinne, Verluste "
                  "wirken sofort. Übertrag aus dem Blatt Neuobjekte, der Rest wird im Fristjahr "
                  "aufgelöst. Steuer hier nur auf Veräußerung und Auflösung, ein Verlust "
                  "mindert sie. " + HINWEIS_FINANZIERUNG)
    ws.freeze_panes = "B2"
    ws.protection.sheet = True


def _jahresblatt(wb, modell: Modell, blatt: str, spalten, zeilen_formeln,
                 hinweis: str) -> None:
    """Rechenblatt mit einer Zeile je Prognosejahr und einer Summenzeile darunter."""
    ws = wb.create_sheet(blatt)
    erste, letzte = 2, prognosejahre(modell) + 1
    buchstaben = formeln.jahres_spalten(spalten)
    _kopf(ws, 1, [s.ueberschrift for s in spalten])
    ws.row_dimensions[1].height = 45
    _rechenspalten(wb, ws, blatt, spalten, buchstaben, "", erste, letzte, zeilen_formeln)

    # Summe über alle Jahre nur für Stromgrößen, nicht für Bestände und kumulierte Werte
    summe = letzte + 1
    for s in spalten:
        bst = buchstaben[s.key]
        c = ws[f"{bst}{summe}"]
        if s.key == "jahr":
            c.value = "Summe"
        elif s.summe:
            c.value = f"=SUM({bst}{erste}:{bst}{letzte})"
            c.number_format = s.format
        c.font = Font(bold=True)
        c.fill = FILL_BERECHNET

    ws.cell(row=1, column=len(spalten) + 2, value=hinweis)
    ws.freeze_panes = "B2"
    ws.protection.sheet = True


def _blatt_liquiditaet(wb, modell: Modell) -> None:
    _jahresblatt(wb, modell, "Liquidität", LIQUIDITAET_SPALTEN, formeln.liquiditaet_zeile,
                 "Summen über alle Objekte und Neuobjekte je Jahr. Freier Mittelzufluss = "
                 "laufender Überschuss + Verkaufserlös netto − Steuer gesamt − Kauf Neuobjekte; "
                 "die AfA fließt nicht ab. Steuer ohne Verlustvortrag: ein Verlust mindert die "
                 "Steuer im selben Jahr. Neuobjekte voll aus Eigenmitteln. "
                 + HINWEIS_FINANZIERUNG)


def _blatt_auswertung(wb, modell: Modell) -> None:
    _jahresblatt(wb, modell, "Auswertung", AUSWERTUNG_SPALTEN, formeln.auswertung_zeile,
                 "Gesamt-GuV = laufendes Ergebnis + steuerpflichtiger Teil aus Verkauf und "
                 "Auflösung (Rücklagenspiegel). Buch- und Verkehrswert über die Objekte im "
                 "Bestand am Jahresende; ohne Verkehrswert im Objektblatt gilt der Buchwert. "
                 "Stille Reserven = Verkehrswert − Buchwert, eine Steuerungsgröße, keine "
                 "Steuerposition. Alternativanlage: freie Mittel zum Jahresende angelegt, "
                 "verzinst ab dem Folgejahr mit der Alternativrendite, Zins versteuert; ein "
                 "negativer Stand wird ebenso verzinst. Vermögen nach Steuern = Verkehrswert + "
                 "Alternativanlage − latente Steuer auf stille Reserven und Rücklage (ohne "
                 "Gewinnzuschlag). " + HINWEIS_FINANZIERUNG)


def _blatt_vergleich(wb, modell: Modell) -> None:
    """Szenariovergleich: Kennzahlen des aktiven Szenarios und gespeicherte Läufe A, B, C.

    Spalten: A Kennzahl, B aktives Szenario, C–E gespeichert A, B, C (Eingabe),
    F Differenz A − B, G Differenz A − C, H Erläuterung.
    """
    ws = wb.create_sheet("Vergleich")
    ws["A1"] = "Szenariovergleich"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = HINWEIS_FINANZIERUNG
    ws["A3"] = "aktives Szenario"
    ws["B3"] = "=par_Szenario"
    ws["B3"].fill = FILL_BERECHNET
    gespeichert_bst = dict(zip(SZENARIEN, "CDE"))
    vergleiche = [(sz, bst) for sz, bst in zip(SZENARIEN[1:], "FG")]
    _kopf(ws, 5, ["Kennzahl", "aktives Szenario"]
          + [f"Szenario {sz} gespeichert" for sz in SZENARIEN]
          + [f"Differenz {SZENARIO_A} − {sz}" for sz, _ in vergleiche] + ["Erläuterung"])
    ws.row_dimensions[5].height = 32
    erste = 6
    letzte = erste + len(VERGLEICH_KENNZAHLEN) - 1
    aktuell = formeln.vergleich_aktuell(erste)
    a = gespeichert_bst[SZENARIO_A]
    for zeile, k in enumerate(VERGLEICH_KENNZAHLEN, start=erste):
        ws.cell(row=zeile, column=1, value=k.bezeichnung)
        zellen = {"B": aktuell[k.key]}
        zellen.update({bst: modell.vergleich.get(sz, {}).get(k.key)
                       for sz, bst in gespeichert_bst.items()})
        for sz, bst in vergleiche:
            andere = f"{gespeichert_bst[sz]}{zeile}"
            zellen[bst] = f'=IF(OR({a}{zeile}="",{andere}=""),"",{a}{zeile}-{andere})'
        for bst, wert in zellen.items():
            c = ws[f"{bst}{zeile}"]
            c.value = wert
            c.number_format = FMT_EURO
            eingabe = bst in gespeichert_bst.values()
            c.fill = FILL_EINGABE if eingabe else FILL_BERECHNET
            if eingabe:
                c.protection = Protection(locked=False)
            if k.fett:
                c.font = Font(bold=True)
        ws[f"H{zeile}"] = k.erlaeuterung
        if k.fett:
            ws[f"A{zeile}"].font = Font(bold=True)
    namen = [("vg_Aktuell", "B")] + [(f"vg_{sz}", bst) for sz, bst in gespeichert_bst.items()]
    namen += [(f"vg_Diff{sz}", bst) for sz, bst in vergleiche]
    for name, bst in namen:
        _name(wb, name, f"Vergleich!${bst}${erste}:${bst}${letzte}")
    # Endvermögen aktiv, gespeichert A, B, C, Differenzen: für die Meldungen der Makros
    z = formeln.vergleich_zeilen(erste)["vermoegen"]
    _name(wb, "vg_Vermoegen", f"Vergleich!$B${z}:$G${z}")

    hinweise = [
        "So wird verglichen: Schaltfläche „Szenarien A, B, C vergleichen“ auf dem "
        "Parameterblatt (Makro modSzenario.SzenarienVergleichenStarten). Es rechnet nacheinander "
        "A, B und C, speichert die Spalte „aktives Szenario“ jeweils als Werte in Spalte C, D "
        "bzw. E und stellt danach das vorher aktive Szenario wieder ein. Ohne Makros: Szenario "
        "auf dem Parameterblatt wählen und die Spalte von Hand als Werte einfügen.",
        "A rechnet die § 6b-Kette wie erfasst. B versteuert jeden Veräußerungsgewinn sofort; "
        "Neuobjekte mit Quelle-Rücklage entfallen, das Kapital bleibt in der Alternativanlage. "
        "C versteuert ebenfalls sofort, kauft die Neuobjekte aber trotzdem, ohne Übertrag und "
        "mit voller AfA-Basis. Neuobjekte ohne Rücklage gibt es in allen Szenarien.",
        "A − C zeigt die reine Wirkung von § 6b: Die Steuer wird nicht gespart, sondern "
        "gestundet und über geringere AfA und höhere latente Steuer nachgeholt. Bei "
        "Alternativrendite null und gleichem Steuersatz ist A − C beim Endvermögen null, "
        "der Vorteil von A ist der Zins auf die gestundete Steuer.",
        "A − B vergleicht dagegen Neuobjekt gegen Alternativanlage: Miete und Wertsteigerung "
        "gegen Zinsertrag, dazu die gestundete Steuer.",
        "In allen Szenarien liegen die freien Mittel in der Alternativanlage. Entscheidend ist "
        "das Endvermögen nach Steuern, nicht die gesparte Steuer allein; die latente Steuer "
        "auf stille Reserven und Restrücklage ist abgezogen, als würde am Ende alles verkauft.",
    ]
    for i, text in enumerate(hinweise, start=letzte + 2):
        ws.cell(row=i, column=1, value=text)

    for spalte, breite in zip("ABCDEFGH", (40, 18, 18, 18, 18, 18, 18, 70)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "B6"
    ws.protection.sheet = True


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
    ergebnis = f"E{erste}:E{letzte}"
    for art, fill in ((FEHLER, FILL_FEHLER), (WARNUNG, FILL_WARNUNG)):
        ws.conditional_formatting.add(
            ergebnis, FormulaRule(formula=[f'E{erste}="{art}"'], fill=fill))

    hinweise = [
        f"{FEHLER}: Eingabe unvollständig oder unzulässig; die Zeile rechnet nicht mit bzw. das "
        "Ergebnis gilt nicht. "
        f"{WARNUNG}: rechnet, aber steuerlich ungünstig oder fachlich zu prüfen. "
        f"{HINWEIS}: zur Kenntnis, zählt nicht ins Gesamtergebnis.",
        "Gesamtergebnis auch auf dem Parameterblatt (Plausibilitätsprüfung). Das Makro "
        "„Plausibilität prüfen“ rechnet neu und listet die auffälligen Prüfungen auf; vor dem "
        "Szenariovergleich fragt es bei Fehlern nach.",
        "Je Objekt, Verkauf und Neuobjekt steht der Grund in der Statusspalte des Blatts. "
        "Die Anteile G+B und Gebäude eines Neuobjekts ergeben durch den Aufbau immer den "
        "Kaufpreis (Anteil G+B, Rest Gebäude).",
    ]
    for i, text in enumerate(hinweise, start=letzte + 2):
        ws.cell(row=i, column=1, value=text)
    for spalte, breite in zip("ABCDEF", (16, 70, 10, 10, 10, 50)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A8"
    ws.protection.sheet = True


def erstelle_mappe(modell: Modell) -> Workbook:
    wb = Workbook()
    _blatt_parameter(wb, modell)
    _blatt_objekte(wb, modell)
    _blatt_verkaeufe(wb, modell)
    _blatt_neuobjekte(wb, modell)
    _blatt_prognose(wb, modell)
    _blatt_ruecklagen(wb, modell)
    _blatt_liquiditaet(wb, modell)
    _blatt_auswertung(wb, modell)
    _blatt_vergleich(wb, modell)
    _blatt_pruefung(wb)
    wb.code_name = CODENAME_MAPPE
    for ws in wb.worksheets:
        ws.sheet_properties.codeName = CODENAMEN[ws.title]
    return wb
