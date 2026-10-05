"""Formel-Bausteine.

Formeln werden in englischer Excel-Syntax geschrieben (so legt openpyxl sie ab);
Excel zeigt sie in der deutschen Oberfläche automatisch als WENN, ZÄHLENWENN usw.
"""

from openpyxl.utils import get_column_letter

from .modelle import OBJEKT_FELDER, PROGNOSE_SPALTEN


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


def _indexiert(name: str, satz: str, zeile: int, objekt_nr: int) -> str:
    return (f"{_stamm(name, objekt_nr)}*(1+{satz})^({_p('jahr', zeile)}-par_Basisjahr)"
            f"*{_p('aktiv', zeile)}")


def prognose_zeile(zeile: int, objekt_nr: int, erstes_jahr: bool) -> dict:
    """Formeln einer Prognosezeile je Spaltenschlüssel.

    objekt_nr ist die Zeile im Objektblatt (1 = erstes Objekt). Leere Objektzeilen
    ergeben leere Prognosezeilen.
    """
    id_obj = _stamm("obj_ID", objekt_nr)
    bw_vor = (_stamm("obj_Restbuchwert", objekt_nr) if erstes_jahr
              else _p("buchwert", zeile - 1))
    afa_voll = f"{_stamm('obj_AKGebaeude', objekt_nr)}*{_stamm('obj_AfASatz', objekt_nr)}"

    def indexiert(name, satz):
        return _leer_oder(zeile, _indexiert(name, satz, zeile, objekt_nr))

    return {
        "id": f'=IF({id_obj}="","",{id_obj})',
        "jahr": "=par_Startjahr" if erstes_jahr else f"={_p('jahr', zeile - 1)}+1",
        # Etappe 4: 0 ab dem Jahr nach dem Verkauf
        "aktiv": _leer_oder(zeile, "1"),
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
    }
