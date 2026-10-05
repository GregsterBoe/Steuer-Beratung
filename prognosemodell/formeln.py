"""Formel-Bausteine.

Formeln werden in englischer Excel-Syntax geschrieben (so legt openpyxl sie ab);
Excel zeigt sie in der deutschen Oberfläche automatisch als WENN, ZÄHLENWENN usw.
"""

from openpyxl.utils import get_column_letter

from .modelle import OBJEKT_FELDER, PROGNOSE_SPALTEN, VERKAUF_FELDER, VERKAUF_SPALTEN


def spalte(key: str) -> str:
    """Spaltenbuchstabe eines Objektfelds."""
    for i, f in enumerate(OBJEKT_FELDER, start=1):
        if f.key == key:
            return get_column_letter(i)
    raise KeyError(key)


def status_objekt(zeile: int) -> str:
    """Plausibilitätsstatus je Objektzeile; leer, solange keine ObjektID steht."""
    id_ = f"${spalte('objekt_id')}{zeile}"
    pflicht = [f"${spalte(f.key)}{zeile}" for f in OBJEKT_FELDER if f.pflicht]
    kaufjahr = f"${spalte('kaufjahr')}{zeile}"
    restbw = f"${spalte('restbuchwert')}{zeile}"
    ak_geb = f"${spalte('ak_gebaeude')}{zeile}"
    return (
        f'=IF({id_}="","",'
        f'IF(COUNTA({",".join(pflicht)})<{len(pflicht)},"Pflichtfeld fehlt",'
        f'IF(COUNTIF(obj_ID,{id_})>1,"ObjektID doppelt",'
        f'IF({kaufjahr}>par_Basisjahr,"Kaufjahr nach Basisjahr",'
        f'IF({restbw}>{ak_geb},"Restbuchwert > AK Gebäude",'
        f'"OK")))))'
    )


# --- Prognosematrix (Etappe 2, Projektplan Abschnitt 8) ---
#
# Jede Zeile des Objektblatts hat einen festen Block aufeinanderfolgender
# Jahreszeilen; die Stammdaten kommen per INDEX aus genau dieser Objektzeile.
# Der Buchwert des Vorjahres steht deshalb direkt in der Zeile darüber; im
# ersten Prognosejahr ist es der Restbuchwert aus dem Objektblatt.


def pspalte(key: str) -> str:
    """Spaltenbuchstabe einer Prognosespalte."""
    for i, s in enumerate(PROGNOSE_SPALTEN, start=1):
        if s.key == key:
            return get_column_letter(i)
    raise KeyError(key)


def _p(key: str, zeile: int) -> str:
    return f"${pspalte(key)}{zeile}"


def _stamm(name: str, objekt_nr: int) -> str:
    """Stammdatum aus Zeile objekt_nr des Objektblatts (1 = erstes Objekt)."""
    return f"INDEX({name},{objekt_nr})"


def _leer_oder(zeile: int, ausdruck: str) -> str:
    return f'=IF({_p("id", zeile)}="","",{ausdruck})'


def _indexiert(name: str, satz: str, zeile: int, objekt_nr: int, aktiv: bool = True) -> str:
    ausdruck = f"{_stamm(name, objekt_nr)}*(1+{satz})^({_p('jahr', zeile)}-par_Basisjahr)"
    return f"{ausdruck}*{_p('aktiv', zeile)}" if aktiv else ausdruck


def _verkaufsjahr(zeile: int) -> str:
    """Verkaufsjahr laut Blatt Verkäufe; ohne Verkauf 9999, also nie."""
    treffer = f"INDEX(vk_Jahr,MATCH({_p('id', zeile)},vk_ID,0))"
    return f'IFERROR(IF({treffer}="",9999,{treffer}),9999)'


def prognose_zeile(zeile: int, objekt_nr: int, erstes_jahr: bool) -> dict:
    """Formeln einer Prognosezeile je Spaltenschlüssel.

    objekt_nr ist die Zeile im Objektblatt (1 = erstes Objekt). Leere Objektzeilen
    ergeben leere Prognosezeilen.
    """
    id_obj = _stamm("obj_ID", objekt_nr)
    bw_vor = (_stamm("obj_Restbuchwert", objekt_nr) if erstes_jahr
              else _p("buchwert", zeile - 1))
    afa_voll = f"{_stamm('obj_AKGebaeude', objekt_nr)}*{_stamm('obj_AfASatz', objekt_nr)}"

    def indexiert(name, satz, aktiv=True):
        return _leer_oder(zeile, _indexiert(name, satz, zeile, objekt_nr, aktiv))

    return {
        "id": f'=IF({id_obj}="","",{id_obj})',
        "jahr": "=par_Startjahr" if erstes_jahr else f"={_p('jahr', zeile - 1)}+1",
        # 0 ab dem Jahr nach dem Verkauf
        "aktiv": _leer_oder(zeile, f"IF({_p('jahr', zeile)}<={_verkaufsjahr(zeile)},1,0)"),
        "miete": indexiert("obj_MieteBasis", "par_Mietsteig"),
        "einnahmen": indexiert("obj_EinnBasis", "par_Mietsteig"),
        "erhaltung": indexiert("obj_ErhBasis", "par_Erhaltsteig"),
        "ausgaben": indexiert("obj_AusgBasis", "par_Kostensteig"),
        # keine AfA über den Restbuchwert hinaus
        "afa": _leer_oder(zeile, f"MIN({afa_voll},{bw_vor})*{_p('aktiv', zeile)}"),
        "buchwert": _leer_oder(zeile, f"MAX({bw_vor}-{_p('afa', zeile)},0)"),
        "ergebnis": _leer_oder(
            zeile,
            f"({_p('miete', zeile)}+{_p('einnahmen', zeile)}-{_p('erhaltung', zeile)}"
            f"-{_p('ausgaben', zeile)}-{_p('afa', zeile)})*{_p('aktiv', zeile)}"),
        "verkehrswert": indexiert("obj_Verkehrswert", "par_Wertsteig", aktiv=False),
        # Verkauf zum Jahresende: im Verkaufsjahr schon nicht mehr im Bestand
        "bestand": _leer_oder(zeile, f"IF({_p('jahr', zeile)}<{_verkaufsjahr(zeile)},1,0)"),
    }


# --- Verkäufe (Etappe 4, Projektplan Abschnitt 11) ---
#
# Verkauf zum Jahresende: Der Gebäudebuchwert ist der Prognosewert am Ende des
# Verkaufsjahrs, also nach dessen AfA. Den Erlös teilt der Anteil G+B laut
# Kaufvertrag auf, ersatzweise der Verkehrswertanteil aus dem Objektblatt.


def vspalte(key: str) -> str:
    """Spaltenbuchstabe im Blatt Verkäufe (Eingaben, dann berechnete Spalten)."""
    for i, s in enumerate(VERKAUF_FELDER + VERKAUF_SPALTEN, start=1):
        if s.key == key:
            return get_column_letter(i)
    raise KeyError(key)


def _v(key: str, zeile: int) -> str:
    return f"${vspalte(key)}{zeile}"


def verkauf_zeile(zeile: int) -> dict:
    """Formeln der berechneten Verkaufsspalten; leer ohne ObjektID oder bei unbekannter ID."""
    id_ = _v("objekt_id", zeile)

    def obj(name):
        return f"INDEX({name},MATCH({id_},obj_ID,0))"

    def wenn(ausdruck):
        return f'=IF(OR({id_}="",COUNTIF(obj_ID,{id_})=0),"",{ausdruck})'

    quote_vertrag = _v("anteil_gub", zeile)
    return {
        "vorbesitz": wenn(f"{_v('jahr', zeile)}-{obj('obj_Kaufjahr')}"),
        "buchwert_geb": wenn(f"SUMIFS(prg_Buchwert,prg_ID,{id_},prg_Jahr,{_v('jahr', zeile)})"),
        "ak_gub": wenn(obj("obj_AKGuB")),
        "nettoerloes": wenn(f"{_v('preis', zeile)}-{_v('kosten', zeile)}"),
        # Kaufvertrag vor Verkehrswert; ohne beides leer, der Status meldet es
        "quote_gub": wenn(f'IF({quote_vertrag}<>"",{quote_vertrag},'
                          f'IF({obj("obj_VKQuoteGeb")}="","",1-{obj("obj_VKQuoteGeb")}))'),
        "erloes_gub": wenn(f'IF({_v("quote_gub", zeile)}="","",'
                           f'{_v("nettoerloes", zeile)}*{_v("quote_gub", zeile)})'),
        "erloes_geb": wenn(f'IF({_v("quote_gub", zeile)}="","",'
                           f'{_v("nettoerloes", zeile)}-{_v("erloes_gub", zeile)})'),
        "gewinn_geb": wenn(f'IF({_v("quote_gub", zeile)}="","",'
                           f'{_v("erloes_geb", zeile)}-{_v("buchwert_geb", zeile)})'),
        # G+B wird nicht abgeschrieben: Buchwert = AK
        "gewinn_gub": wenn(f'IF({_v("quote_gub", zeile)}="","",'
                           f'{_v("erloes_gub", zeile)}-{_v("ak_gub", zeile)})'),
        "gewinn": wenn(f'IF({_v("quote_gub", zeile)}="","",'
                       f'{_v("gewinn_geb", zeile)}+{_v("gewinn_gub", zeile)})'),
    }


def status_verkauf(zeile: int) -> str:
    """Plausibilitätsstatus je Verkaufszeile; leer, solange keine ObjektID steht."""
    id_, jahr = _v("objekt_id", zeile), _v("jahr", zeile)
    return (
        f'=IF({id_}="","",'
        f'IF(COUNTIF(obj_ID,{id_})=0,"ObjektID unbekannt",'
        f'IF(COUNTIF(vk_ID,{id_})>1,"Objekt mehrfach verkauft",'
        f'IF({jahr}="","Verkaufsjahr fehlt",'
        f'IF(OR({jahr}<par_Startjahr,{jahr}>par_Endjahr),"Verkaufsjahr außerhalb Raster",'
        f'IF({_v("preis", zeile)}="","Verkaufspreis fehlt",'
        f'IF({_v("quote_gub", zeile)}="","Aufteilung fehlt: Anteil G+B oder Verkehrswertanteil",'
        f'IF(AND({_v("nutzung_6b", zeile)}="ja",{_v("vorbesitz", zeile)}<par_6bVorbesitz),'
        f'"§ 6b unzulässig: Vorbesitzzeit zu kurz",'
        f'"OK"))))))))'
    )
