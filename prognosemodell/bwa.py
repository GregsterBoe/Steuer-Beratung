"""Ausgabe im DATEV-BWA-Format und Sonderbereich Verkauf und Kauf (Projektplan Abschnitt 18).

Je Kostenstelle ein Blatt im Layout der Kostenstellenblätter (vorlagen.bwa_kopf):
links die Ist-Werte, rechts die Planjahre als Formeln aus der Prognose (Szenario A).
Das Summenblatt „Alle Objekte“ rechnet über alle Objekte, auch über solche, die
erst in Excel angelegt werden und kein eigenes Blatt haben, und stimmt mit dem
Blatt Liquidität überein: Ergebnis vor Steuern = Ergebnis vor Verlustvortrag.

Das Blatt „Verkauf und Kauf“ zeigt je Verkauf eine Ergebnissicht und eine
Detailsicht, darunter je Neuobjekt die Detailsicht des Kaufs.
"""

import re

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .modelle import (FMT_EURO, FMT_JAHR, FMT_PROZENT, MAX_NEUOBJEKTE, MAX_VERKAEUFE,
                      PARAMETER, STATUS_6B_UNZULAESSIG, STATUS_OK, Modell, prognosejahre)
from .vorlagen import (BWA_ZEILEN, ERSTE_ZEILE, SPALTE_JAHR, SPALTE_PLAN, SPALTE_VORJAHRE,
                       bwa_kopf)

SUMMENBLATT = "Alle Objekte"
SONDERBEREICH = "Verkauf und Kauf"
ZEILE_JAHR = ERSTE_ZEILE - 1   # Hilfszeile mit dem Planjahr als Zahl, für die Formeln

# Kostenarten aus „weitere Ausgaben“; 1260 Sonstige Kosten nimmt den Rest auf
AUSGABEN_ZEILEN = (1100, 1120, 1140, 1150, 1180, 1200, 1220)
KOSTENARTEN = AUSGABEN_ZEILEN + (1240, 1250, 1260)
# Ist-Werte im Basisjahr aus dem Objektblatt, solange kein Kostenstellenblatt eingelesen ist
IST_AUS_OBJEKT = {1020: "obj_MieteBasis", 1090: "obj_EinnBasis", 1250: "obj_ErhBasis",
                  1260: "obj_AusgBasis"}
# mindestens so viele Vorgänge und Neuobjekte im Sonderbereich, damit in Excel ergänzte passen
MIN_VORGAENGE = 3

FMT_BWA = '#,##0.00;-#,##0.00;""'   # 0 bleibt leer, etwa nach dem Verkauf
FONT_TITEL = Font(bold=True, size=14)
FONT_ABSCHNITT = Font(bold=True, size=12)
FONT_HILFE = Font(italic=True, color="808080", size=8)
FILL_BERECHNET = PatternFill("solid", fgColor="E7E6E6")
FILL_KOPF = PatternFill("solid", fgColor="D9E1F2")
GUELTIG = (STATUS_OK, STATUS_6B_UNZULAESSIG)

UNGUELTIGE_ZEICHEN = re.compile(r"[\[\]:*?/\\]")


def _basisjahr(modell: Modell) -> int:
    return modell.parameter.get(
        "par_Basisjahr", next(p.wert for p in PARAMETER if p.name == "par_Basisjahr"))


def blattname(wunsch: str, vorhanden) -> str:
    """Gültiger, eindeutiger Blattname (höchstens 31 Zeichen, ohne []:*?/\\)."""
    basis = UNGUELTIGE_ZEICHEN.sub("-", str(wunsch)).strip("'")[:31] or "BWA"
    name, i = basis, 2
    belegt = {t.lower() for t in vorhanden}
    while name.lower() in belegt:
        name = f"BWA {basis}"[:31] if i == 2 else f"{basis[:27]} ({i})"
        i += 1
    return name


def _summenformeln(z: dict, sp: str) -> dict:
    """Summenzeilen der BWA Form 01 als Formeln, wie im Kostenstellenblatt der Kanzlei."""
    def c(nr):
        return f"{sp}{z[nr]}"

    return {
        1051: f"={c(1020)}",
        1080: f"={c(1051)}",
        1092: f"={c(1080)}+{c(1090)}",
        1280: "=" + "+".join(c(nr) for nr in KOSTENARTEN),
        1300: f"={c(1092)}-{c(1280)}",
        1320: f"={c(1310)}+{c(1312)}",
        1330: f"={c(1322)}+{c(1323)}",
        1345: f"={c(1300)}-{c(1320)}+{c(1330)}",
        1353: f"={c(1345)}+{c(1351)}-{c(1352)}",
        1380: f"={c(1353)}-{c(1355)}",
    }


def _gewinne(jahr: str, vergleich: str, *bedingung: str) -> str:
    """Veräußerungsgewinne gültiger Verkäufe im Jahr; vergleich ">0" oder "<0"."""
    extra = "".join(f",{b}" for b in bedingung)
    return "+".join(f'SUMIFS(vk_Gewinn,vk_Jahr,{jahr},vk_Status,"{st}",vk_Gewinn,"{vergleich}"{extra})'
                    for st in GUELTIG)


def _planformeln_objekt(z: dict, sp: str, neu: bool) -> dict:
    """Planspalte eines Kostenstellenblatts; ObjektID in $B$2, Planjahr in Zeile 5."""
    jahr = f"{sp}${ZEILE_JAHR}"

    def p(name):
        return f"SUMIFS({name},prg_ID,$B$2,prg_Jahr,{jahr})"

    formeln = {
        1020: f"={p('prg_Miete')}",
        1090: f"={p('prg_Einnahmen')}",
        1240: f"={p('prg_AfA')}",
        1250: f"={p('prg_Erhaltung')}",
    }
    if neu:
        formeln[1260] = f"={p('prg_Ausgaben')}"
    else:
        # Aufteilung der weiteren Ausgaben nach dem Anteil der Kostenart im Basisjahr
        basis = "IFERROR(INDEX(obj_AusgBasis,MATCH($B$2,obj_ID,0)),0)"
        for nr in AUSGABEN_ZEILEN:
            formeln[nr] = (f"=IF(N({basis})=0,0,{p('prg_Ausgaben')}"
                           f"*N(${get_column_letter(SPALTE_JAHR)}${z[nr]})/{basis})")
        andere = "+".join(f"{sp}{z[nr]}" for nr in AUSGABEN_ZEILEN)
        formeln[1260] = f"={p('prg_Ausgaben')}-({andere})"
        # Verkauf: Gewinn als neutraler Ertrag, Verlust und Einstellung in die Rücklage
        # als neutraler Aufwand; Auflösung und Zuschlag im Fristjahr als Ertrag
        objekt = "vk_ID,$B$2"
        formeln[1323] = (f"={_gewinne(jahr, '>0', objekt)}"
                         f"+SUMIFS(rl_Aufloesung,rl_ObjektID,$B$2,rl_Fristjahr,{jahr})"
                         f"+SUMIFS(rl_Zuschlag,rl_ObjektID,$B$2,rl_Fristjahr,{jahr})")
        formeln[1312] = (f"=-({_gewinne(jahr, '<0', objekt)})"
                         f"+SUMIFS(rl_Betrag,rl_ObjektID,$B$2,rl_Jahr,{jahr})")
    formeln.update(_summenformeln(z, sp))
    return formeln


def _planformeln_summe(z: dict, sp: str, blaetter: list) -> dict:
    """Planspalte des Summenblatts: Szenario A über alle Objekte, auch ohne eigenes Blatt."""
    jahr = f"{sp}${ZEILE_JAHR}"

    def p(name):
        return f"SUMIFS({name},prg_Jahr,{jahr})"

    zins = f"SUMIFS(liq_Zins,liq_Jahr,{jahr})"
    formeln = {
        1020: f"={p('prg_Miete')}",
        1090: f"={p('prg_Einnahmen')}",
        1240: f"={p('prg_AfA')}",
        1250: f"={p('prg_Erhaltung')}",
        # Zins der Alternativanlage; negative Liquidität kostet Zins
        1310: f"=MAX(-{zins},0)",
        1322: f"=MAX({zins},0)",
        1323: (f"={_gewinne(jahr, '>0')}+SUMIFS(rls_Aufloesung,rls_Jahr,{jahr})"
               f"+SUMIFS(rls_Zuschlag,rls_Jahr,{jahr})"),
        1312: f"=-({_gewinne(jahr, '<0')})+SUMIFS(rls_Bildung,rls_Jahr,{jahr})",
        1355: f"=SUMIFS(liq_Steuer,liq_Jahr,{jahr})",
    }
    for nr in AUSGABEN_ZEILEN:
        formeln[nr] = "=" + ("+".join(f"'{b}'!{sp}{z[nr]}" for b in blaetter) or "0")
    andere = "+".join(f"{sp}{z[nr]}" for nr in AUSGABEN_ZEILEN)
    formeln[1260] = f"={p('prg_Ausgaben')}-({andere})"
    formeln.update(_summenformeln(z, sp))
    return formeln


def _jahreszeile(ws, planjahre: int) -> None:
    ws.cell(row=ZEILE_JAHR, column=3, value="Jahr (für Formeln)").font = FONT_HILFE
    ws.cell(row=ZEILE_JAHR, column=SPALTE_JAHR, value="=par_Basisjahr").font = FONT_HILFE
    for i in range(planjahre):
        c = ws.cell(row=ZEILE_JAHR, column=SPALTE_PLAN + i, value=f"=par_Startjahr+{i}")
        c.font = FONT_HILFE
        c.number_format = FMT_JAHR


def _planspalten(ws, z: dict, planjahre: int, formeln_je_spalte) -> None:
    for i in range(planjahre):
        sp = get_column_letter(SPALTE_PLAN + i)
        for nr, formel in formeln_je_spalte(sp).items():
            c = ws.cell(row=z[nr], column=SPALTE_PLAN + i, value=formel)
            c.number_format = FMT_BWA
            c.fill = FILL_BERECHNET


def _blatt_objekt(wb, titel: str, objekt_id: str, name, ist, basisjahr: int,
                  planjahre: int, neu: bool):
    ws = wb.create_sheet(titel)
    z = bwa_kopf(ws, objekt_id, name or "", basisjahr, planjahre)
    _jahreszeile(ws, planjahre)
    if ist:
        for nr, werte in ist.items():
            if nr in z:
                for spalte, wert in werte.items():
                    ws.cell(row=z[nr], column=spalte, value=wert)
    elif not neu:
        # ohne eingelesene BWA: Basisjahr aus dem Objektblatt, Summenzeilen als Formel
        t = get_column_letter(SPALTE_JAHR)
        for nr, feld in IST_AUS_OBJEKT.items():
            ws.cell(row=z[nr], column=SPALTE_JAHR,
                    value=f"=IFERROR(N(INDEX({feld},MATCH($B$2,obj_ID,0))),0)")
        for nr, formel in _summenformeln(z, t).items():
            ws.cell(row=z[nr], column=SPALTE_JAHR, value=formel)
    _planspalten(ws, z, planjahre, lambda sp: _planformeln_objekt(z, sp, neu))
    hinweis = ("Neuobjekt: Werte ab dem Jahr nach dem Kauf. " if neu else "") + (
        "Planspalten = Szenario A, vor Finanzierung; Zins und Steuer nur im Blatt "
        f"„{SUMMENBLATT}“.")
    ws["E2"] = hinweis
    ws["E2"].font = Font(italic=True)
    return ws


def blaetter_bwa(wb, modell: Modell) -> list:
    """Summenblatt und je Objekt und Neuobjekt ein BWA-Blatt; liefert die Blattnamen."""
    basisjahr = _basisjahr(modell)
    planjahre = prognosejahre()
    summe = wb.create_sheet(blattname(SUMMENBLATT, wb.sheetnames))
    eintraege = []
    for obj in modell.objekte:
        if not obj.objekt_id:
            continue
        ist = modell.kostenstellen.get(obj.objekt_id)
        wunsch = ist.blatt if ist is not None else obj.objekt_id
        eintraege.append((wunsch, obj.objekt_id, obj.name, ist.ist if ist else None, False))
    for ne in modell.neuobjekte:
        if ne.neu_id:
            eintraege.append((ne.neu_id, ne.neu_id, ne.name, None, True))
    titel = []
    for wunsch, objekt_id, name, ist, neu in eintraege:
        t = blattname(wunsch, wb.sheetnames)
        _blatt_objekt(wb, t, objekt_id, name, ist, basisjahr, planjahre, neu)
        titel.append(t)

    z = bwa_kopf(summe, "KSt", SUMMENBLATT, basisjahr, planjahre)
    _jahreszeile(summe, planjahre)
    bestand = [t for t, e in zip(titel, eintraege) if not e[4]]
    # Ist-Werte: Summe der Kostenstellenblätter
    for spalte in range(SPALTE_VORJAHRE, SPALTE_JAHR + 1):
        sp = get_column_letter(spalte)
        for nr, bez in BWA_ZEILEN:
            if bez and bestand:
                summe.cell(row=z[nr], column=spalte,
                           value="=" + "+".join(f"N('{b}'!{sp}{z[nr]})" for b in bestand))
    _planspalten(summe, z, planjahre, lambda sp: _planformeln_summe(z, sp, titel))
    summe["E2"] = ("Planspalten = Szenario A über alle Objekte und Neuobjekte, vor "
                   "Finanzierung. Ergebnis vor Steuern = Ergebnis vor Verlustvortrag im Blatt "
                   "Liquidität, Steuer mit Verlustvortrag.")
    summe["E2"].font = Font(italic=True)
    ende = get_column_letter(SPALTE_PLAN + planjahre - 1)
    anfang = get_column_letter(SPALTE_PLAN)
    for nr, bez in BWA_ZEILEN:
        if bez:
            _name(wb, f"bwa_{nr}", f"'{summe.title}'!${anfang}${z[nr]}:${ende}${z[nr]}")
    _name(wb, "bwa_Jahr", f"'{summe.title}'!${anfang}${ZEILE_JAHR}:${ende}${ZEILE_JAHR}")
    return [summe.title] + titel


def _name(wb, name: str, ref: str) -> None:
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# --- Sonderbereich Verkauf und Kauf ---


def _text(zelle, text: str):
    """Beschriftung als Text, auch wenn sie mit „=“ beginnt wie in der Kanzleivorlage."""
    zelle.value = text
    if isinstance(text, str) and text.startswith("="):
        zelle.data_type = "s"
    return zelle


class _Block:
    """Schreibt Zeilen eines Blocks; Formeln nennen andere Zeilen über {schlüssel}."""

    def __init__(self, ws, zeile: int):
        self.ws, self.zeile, self.adr = ws, zeile, {}

    def leer(self, n: int = 1) -> None:
        self.zeile += n

    def titel(self, text, font=FONT_ABSCHNITT) -> None:
        self.ws.cell(row=self.zeile, column=1, value=text).font = font
        self.zeile += 1

    def kopf(self, *texte) -> None:
        for spalte, text in enumerate(texte, start=1):
            c = self.ws.cell(row=self.zeile, column=spalte, value=self._ersetzen(text))
            c.font = Font(bold=True)
            c.fill = FILL_KOPF
            c.alignment = Alignment(wrap_text=True)
        self.zeile += 1

    def zeile_(self, key, text, *formeln, fmt=FMT_EURO, fett=False) -> None:
        """Zeile mit Beschriftung in A und Formeln ab B; {key} = Zelle B, {key.C} = Zelle C."""
        if key:
            self.adr[key] = self.zeile
        _text(self.ws.cell(row=self.zeile, column=1), text).font = Font(bold=fett)
        for spalte, formel in enumerate(formeln, start=2):
            if formel is None:
                continue
            c = self.ws.cell(row=self.zeile, column=spalte, value=self._ersetzen(formel))
            c.number_format = fmt
            c.font = Font(bold=fett)
            if isinstance(formel, str) and formel.startswith("="):
                c.fill = FILL_BERECHNET
        self.zeile += 1

    def _ersetzen(self, formel):
        if not isinstance(formel, str):
            return formel

        def adresse(m):
            key, _, spalte = m.group(1).partition(".")
            return f"${spalte or 'B'}${self.adr[key]}"
        return re.sub(r"\{(\w+(?:\.[A-Z])?)\}", adresse, formel)


def _vorgang(b: _Block, n: int) -> None:
    """Ergebnis- und Detailsicht eines Verkaufs (Zeile n im Blatt Verkäufe)."""
    def vk(name):
        return f"INDEX({name},{n})"

    def rl(name):
        return f"N(INDEX({name},{n}))"

    def ok(ausdruck):
        return f'=IF({{ok}}=1,{ausdruck},"")'

    def halten(feld, satz):
        return f"N(INDEX({feld},MATCH({{id}},obj_ID,0)))*(1+{satz})^({{x}}-par_Basisjahr)"

    def alt(name):
        return f'IF({{rl}}="",0,SUMIFS({name},prg_Quelle,{{rl}},prg_Jahr,{{x}}))'

    def bw_halten(jahr):
        return f"SUMIFS(prg_BuchwertHalten,prg_ID,{{id}},prg_Jahr,{jahr})"

    status = vk("vk_Status")
    b.titel(f'=IF({vk("vk_ID")}="","Vorgang {n}: kein Verkauf erfasst","Vorgang {n}: Verkauf "'
            f'&{vk("vk_ID")}&IFERROR(" "&INDEX(obj_Name,MATCH({vk("vk_ID")},obj_ID,0)),"")'
            f'&" zum Ende "&{vk("vk_Jahr")})')
    b.zeile_("id", "Kostenstelle (ObjektID)", f'=IF({vk("vk_ID")}="","",{vk("vk_ID")})',
             fmt="@")
    b.zeile_("status", "Status Verkauf", f'=IF({{id}}="","",{status})', fmt="@")
    b.zeile_("ok", "rechnet mit (1 = ja)",
             f'=IF(OR({{status}}="{GUELTIG[0]}",{{status}}="{GUELTIG[1]}"),1,0)', fmt="0")
    b.zeile_("jahr", "Verkaufsjahr (Verkauf zum Jahresende)", ok(vk("vk_Jahr")), fmt=FMT_JAHR)
    b.zeile_("rl", "RücklageID (§ 6b)", ok(f'IF({vk("rl_ID")}="","",{vk("rl_ID")})'), fmt="@")
    # Kaufjahr&"0" / 10: leere Annahmeformeln ("") zählen als 0 statt #WERT; geht in
    # jeder Excel-Version ohne Matrixformel (MAXIFS erst ab Excel 2019)
    letzter_kauf = (f'IF({{rl}}="",0,SUMPRODUCT(MAX((ne_Quelle={{rl}})'
                    f'*(ne_Status="{STATUS_OK}")*(ne_Kaufjahr&"0")/10)))')
    b.zeile_("x", "Vergleichsjahr: erstes volles Jahr nach Verkauf und Kauf",
             ok(f"MIN(MAX({{jahr}},{letzter_kauf})+1,par_Endjahr)"), fmt=FMT_JAHR)
    b.leer()

    # Planung im Vergleichsjahr (unten); die Ergebnissicht fasst sie zusammen
    b.titel("Ergebnissicht (für Berichterstattung)")
    b.zeile_("preis", "Bewertung/Erlös (Verkaufspreis)", ok(f"N({vk('vk_Preis')})"))
    reinvest = "+".join(f'SUMIFS({f},ne_Quelle,{{rl}},ne_Status,"{STATUS_OK}")'
                        for f in ("ne_Kaufpreis", "ne_Nebenkosten"))
    b.zeile_("reinvest", "Reinvestition (Kaufpreis und Nebenkosten der Neuobjekte)",
             ok(f'IF({{rl}}="",0,{reinvest})'))
    b.zeile_("uebertrag", "Übertrag § 6b EStG", ok(f"-({rl('rl_UebGeb')}+{rl('rl_UebGuB')})"))
    b.zeile_("steuer", "Steuer auf den Gewinn ca. (Verkaufsjahr, Auflösung im Fristjahr)",
             ok(f"-(N({vk('vk_Gewinn')})-{rl('rl_Betrag')}+{rl('rl_Aufloesung')}"
                f"+{rl('rl_Zuschlag')})*par_Steuersatz"))
    b.zeile_("anlage", "Kapitalanlage (Nettoerlös − Reinvestition − Steuer ca.)",
             ok(f"N({vk('vk_Nettoerloes')})-{{reinvest}}+{{steuer}}"))
    b.zeile_(None, "Restschuld aktuell / nach Umstrukturierung", "Stufe 2 (Finanzierung)",
             fmt="@")
    b.leer()
    b.kopf(f'=IF({{ok}}=1,"Vergleich im Jahr "&{{x}},"Vergleich")', "Ausgangsfall (halten)",
           "Alternative (Verkauf, Kauf, Anlage)", "Differenz")
    zus = [("e_miete", "Mietertrag (netto)", "{p_miete}+{p_einn}"),
           ("e_kapital", "Kapitalertrag", "{p_zins}"),
           ("e_aufwand", "Aufwand (Erhaltung, weitere Ausgaben, AfA; ohne Zinsen)",
            "{p_erh}+{p_ausg}+{p_afa}"),
           ("e_ergebnis", "vorläufiges Ergebnis", "{p_ergebnis}"),
           ("e_vor", "liquider Überschuss vor Steuern", "{p_cashflow}"),
           ("e_steuer", "Steuern ca.", "-{p_ergebnis}*par_Steuersatz"),
           ("e_nach", "liquider Überschuss nach Steuern ca.", "{e_vor}+{e_steuer}")]
    ergebnis_zeile = b.zeile
    b.leer(len(zus) + 1)   # Platz; gefüllt, sobald die Planung steht

    b.titel("Detailsicht (Einzelauflistung)")
    b.kopf("Verkauf", "gesamt", "G+B", "Gebäude")
    q = f"N({vk('vk_AnteilGuB')})"
    b.zeile_("anteil", "Anteil", ok("1"), ok(q), ok(f"1-{q}"), fmt=FMT_PROZENT)
    b.zeile_("vp", "Veräußerungspreis", "={preis}", ok("{preis}*{anteil.C}"),
             ok("{preis}*{anteil.D}"))
    kosten = f"N({vk('vk_Kosten')})"
    b.zeile_("vkosten", "./. Veräußerungskosten", ok(f"-{kosten}"),
             ok(f"-{kosten}*{{anteil.C}}"), ok(f"-{kosten}*{{anteil.D}}"))
    b.zeile_("bw", "./. Buchwert (Gebäude zum Ende des Verkaufsjahrs, G+B = AK)",
             ok("{bw.C}+{bw.D}"), ok(f"-N({vk('vk_AKGuB')})"),
             ok(f"-N({vk('vk_BuchwertGeb')})"))
    b.zeile_("gewinn", "= Veräußerungsgewinn", ok(f"N({vk('vk_Gewinn')})"),
             ok(f"N({vk('vk_GewinnGuB')})"), ok(f"N({vk('vk_GewinnGeb')})"), fett=True)
    b.zeile_("ruecklage", "Rücklage § 6b (gebildet im Verkaufsjahr)", ok(rl("rl_Betrag")),
             ok(rl("rl_GuB")), ok(rl("rl_Geb")))
    b.zeile_("uebertragen", "übertragen auf Neuobjekte (G+B-Gewinn auf G+B und Gebäude)",
             ok("{uebertragen.C}+{uebertragen.D}"), ok(rl("rl_UebGuB")), ok(rl("rl_UebGeb")))
    b.zeile_("frist", "Fristjahr", ok(f'IF({{rl}}="","",{rl("rl_Fristjahr")})'), fmt=FMT_JAHR)
    b.zeile_("aufl", "aufgelöst im Fristjahr", ok(rl("rl_Aufloesung")))
    b.zeile_("zuschlag", "Gewinnzuschlag", ok(rl("rl_Zuschlag")))
    b.leer()

    b.kopf(f'=IF({{ok}}=1,"Planung im Jahr "&{{x}},"Planung")', "Ausgangsfall (halten)",
           "Alternative (Verkauf, Kauf, Anlage)", "Differenz")
    afa_h = f"{bw_halten('{x}-1')}-{bw_halten('{x}')}"
    planung = [
        ("p_miete", "Mieten", halten("obj_MieteBasis", "par_Mietsteig"), alt("prg_Miete")),
        ("p_einn", "weitere Einnahmen", halten("obj_EinnBasis", "par_Mietsteig"),
         alt("prg_Einnahmen")),
        ("p_zins", "Zinsertrag Kapitalanlage", "0", "{anlage}*par_Alternativrendite"),
        ("p_erh", "./. Erhaltung (mit Alterung und Großmaßnahme)",
         "-SUMIFS(prg_ErhaltungHalten,prg_ID,{id},prg_Jahr,{x})",
         "-" + alt("prg_Erhaltung")),
        ("p_ausg", "./. weitere Ausgaben", "-" + halten("obj_AusgBasis", "par_Kostensteig"),
         "-" + alt("prg_Ausgaben")),
        ("p_afa", "./. Abschreibungen (Steuerbilanz)", f"-({afa_h})", "-" + alt("prg_AfA")),
        ("p_zinsen", "./. Zinsen (Stufe 2)", "0", "0"),
        ("p_ergebnis", "= vorläufiges Ergebnis", "{p_miete}+{p_einn}+{p_zins}+{p_erh}+{p_ausg}"
         "+{p_afa}+{p_zinsen}", None),
        ("p_tilgung", "./. Tilgungen (Stufe 2)", "0", "0"),
        ("p_afa_zurueck", "+ Abschreibungen", "-{p_afa}", None),
        ("p_cashflow", "= Cash Flow", "{p_ergebnis}+{p_tilgung}+{p_afa_zurueck}", None),
    ]
    for key, text, ausgang, alternative in planung:
        if alternative is None:   # Summenzeile: gleiche Rechnung in B und C
            alternative = re.sub(r"\{(\w+)\}", r"{\1.C}", ausgang)
        b.zeile_(key, text, ok(ausgang), ok(alternative), ok(f"{{{key}.C}}-{{{key}}}"),
                 fett=text.startswith("="))

    # Ergebnissicht: Zusammenfassung der Planung
    weiter = b.zeile
    b.zeile = ergebnis_zeile
    for key, text, formel in zus:
        alternative = re.sub(r"\{(\w+)\}", r"{\1.C}", formel)
        b.zeile_(key, text, ok(formel), ok(alternative), ok(f"{{{key}.C}}-{{{key}}}"),
                 fett=key in ("e_ergebnis", "e_nach"))
    b.zeile = weiter
    b.leer(2)


def _kauf(b: _Block, anzahl: int) -> None:
    """Detailsicht je Neuobjekt, eine Spalte je Zeile im Blatt Neuobjekte."""
    b.titel("Detailsicht Kauf (je Neuobjekt)")
    b.kopf("Neuobjekt", *[f"Neuobjekt {m}" for m in range(1, anzahl + 1)])

    def zeile(text, ausdruck, fmt=FMT_EURO, fett=False):
        formeln = []
        for m in range(1, anzahl + 1):
            ne_id = f"INDEX(ne_ID,{m})"
            formeln.append(f'=IF({ne_id}="","",{ausdruck.format(m=m, id=ne_id, sp="{sp}")})')
        _text(b.ws.cell(row=b.zeile, column=1), text).font = Font(bold=fett)
        for spalte, formel in enumerate(formeln, start=2):
            formel = formel.replace("{sp}", get_column_letter(spalte))
            c = b.ws.cell(row=b.zeile, column=spalte, value=formel)
            c.number_format = fmt
            c.fill = FILL_BERECHNET
            c.font = Font(bold=fett)
        b.zeile += 1
        return b.zeile - 1

    def ne(name):
        return f"INDEX({name},{{m}})"

    zeile("NeuID (Kostenstelle)", "{id}", "@", fett=True)
    zeile("Name", f'IF({ne("ne_Name")}="","",{ne("ne_Name")})', "@")
    zeile("Status", ne("ne_Status"), "@")
    zeile("Quelle RücklageID", f'IF({ne("ne_Quelle")}="","",{ne("ne_Quelle")})', "@")
    kj = zeile("Kaufjahr (Kauf zum Jahresende)", ne("ne_Kaufjahr"), FMT_JAHR)
    zeile("Kaufpreis", f"N({ne('ne_Kaufpreis')})")
    zeile("Kaufnebenkosten", f"N({ne('ne_Nebenkosten')})")
    zeile("Anteil G+B", f"N({ne('ne_AnteilGuB')})", FMT_PROZENT)
    zeile("AK G+B", f"N({ne('ne_AKGuBNeu')})")
    zeile("AK Gebäude", f"N({ne('ne_AKGebNeu')})")
    zeile("ü1 Gebäudegewinn auf Gebäude", f"-N({ne('ne_Ue1')})")
    zeile("ü2 G+B-Gewinn auf G+B", f"-N({ne('ne_Ue2')})")
    zeile("ü3 G+B-Gewinn auf Gebäude", f"-N({ne('ne_Ue3')})")
    zeile("Übertrag § 6b EStG gesamt", f"-N({ne('ne_UeGesamt')})", fett=True)
    zeile("AfA-Bemessungsgrundlage Gebäude", f"N({ne('ne_AfABasis')})", fett=True)
    zeile("steuerliche AK G+B", f"N({ne('ne_AKGuB')})")
    zeile("AfA-Methode", f'IF({ne("ne_AfAMethode")}="","linear",{ne("ne_AfAMethode")})', "@")
    zeile("AfA-Satz", f"N({ne('ne_AfASatz')})", FMT_PROZENT)
    jahr = zeile("erstes volles Jahr", f"{{sp}}{kj}+1", FMT_JAHR)

    def prg(name):
        return f"SUMIFS({name},prg_ID,{{id}},prg_Jahr,{{sp}}{jahr})"

    miete = zeile("Mieten", prg("prg_Miete"))
    erh = zeile("./. Erhaltung", f"-{prg('prg_Erhaltung')}")
    afa = zeile("./. Abschreibungen (Steuerbilanz)", f"-{prg('prg_AfA')}")
    erg = zeile("= vorläufiges Ergebnis", f"{{sp}}{miete}+{{sp}}{erh}+{{sp}}{afa}", fett=True)
    zeile("= Cash Flow (vor Finanzierung)", f"{{sp}}{erg}-{{sp}}{afa}", fett=True)


def blatt_sonderbereich(wb, modell: Modell) -> None:
    ws = wb.create_sheet(blattname(SONDERBEREICH, wb.sheetnames))
    ws["A1"] = "Sonderbereich Verkauf und Kauf"
    ws["A1"].font = FONT_TITEL
    ws["A2"] = ("Alle Werte vor Finanzierung (ohne Zins und Tilgung), Steuerbilanz. Je Verkauf "
                "eine Ergebnissicht und eine Detailsicht; Steuern ca. = Ergebnis × "
                "Grenzsteuersatz, ohne Verlustvortrag. Die genaue Steuer je Jahr steht im "
                "Blatt Liquidität.")
    b = _Block(ws, 4)
    b.kopf("Endvermögen nach latenter Steuer am Ende des Rasters", "A Plan (§ 6b)",
           "B sofort versteuern", "C versteuern und kaufen", "Baseline halten")
    b.zeile_(None, "Endvermögen", "=INDEX(vg_A,1)", "=INDEX(vg_B,1)", "=INDEX(vg_C,1)",
             "=INDEX(vg_Baseline,1)", fett=True)
    b.zeile_(None, "Differenz zu A", None, "=$B${0}-C{0}".format(b.zeile - 1),
             "=$B${0}-D{0}".format(b.zeile - 1), "=$B${0}-E{0}".format(b.zeile - 1))
    b.leer()

    vorgaenge = min(max(len(modell.verkaeufe), MIN_VORGAENGE), MAX_VERKAEUFE)
    for n in range(1, vorgaenge + 1):
        b.adr = {}
        _vorgang(b, n)
    neu = min(max(len(modell.neuobjekte), MIN_VORGAENGE), MAX_NEUOBJEKTE)
    _kauf(b, neu)
    b.zeile_(None, f"Weitere Vorgänge: das Blatt zeigt {vorgaenge} Verkäufe und {neu} "
             "Neuobjekte, jeweils die ersten Zeilen der Eingabeblätter. Mehr beim nächsten "
             "Generieren.", fmt="@")

    ws.column_dimensions["A"].width = 58
    for spalte in range(2, max(neu, 4) + 2):
        ws.column_dimensions[get_column_letter(spalte)].width = 22
    ws.freeze_panes = "B4"
