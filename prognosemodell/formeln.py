"""Formel-Bausteine.

Formeln werden in englischer Excel-Syntax geschrieben (so legt openpyxl sie ab);
Excel zeigt sie in der deutschen Oberfläche automatisch als WENN, ZÄHLENWENN usw.
"""

from openpyxl.utils import get_column_letter

from .modelle import OBJEKT_FELDER


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


def _obj(key: str, objekt_zeile: int) -> str:
    """Absoluter Verweis auf ein Feld einer Objektzeile."""
    return f"Objekte!${spalte(key)}${objekt_zeile}"


def prognose_zeile(zeile: int, objekt_zeile: int, jahr_index: int) -> dict:
    """Formeln einer Prognosezeile (Spalten A–H laut Projektplan Abschnitt 8).

    Jeder Objektzeile ist ein fester Block von Jahreszeilen zugeordnet, deshalb
    verweist die Zeile direkt auf ihre Objektzeile statt per SVERWEIS über die ID.
    Der Buchwert des Vorjahres steht in der Zeile darüber; im ersten Jahr greift
    der Restbuchwert aus dem Objektblatt.
    """
    status = f"Objekte!${get_column_letter(len(OBJEKT_FELDER) + 1)}${objekt_zeile}"
    bw_vorjahr = _obj("restbuchwert", objekt_zeile) if jahr_index == 0 else f"$G{zeile - 1}"
    afa_voll = f"{_obj('ak_gebaeude', objekt_zeile)}*{_obj('afa_satz', objekt_zeile)}"
    return {
        "id": f'=IF({_obj("objekt_id", objekt_zeile)}="","",{_obj("objekt_id", objekt_zeile)})',
        "jahr": f"=par_Startjahr+{jahr_index}",
        # nur Objekte mit Status OK rechnen; Etappe 4 ergänzt das Verkaufsjahr
        "aktiv": f'=IF({status}="OK",1,0)',
        "miete": f"={_obj('miete', objekt_zeile)}*(1+par_Mietsteig)^($B{zeile}-par_Basisjahr)"
                 f"*$C{zeile}",
        "erhaltung": f"={_obj('erhaltung', objekt_zeile)}"
                     f"*(1+par_Erhaltsteig)^($B{zeile}-par_Basisjahr)*$C{zeile}",
        "afa": f"=MIN({afa_voll},{bw_vorjahr})*$C{zeile}",
        "buchwert": f"=MAX({bw_vorjahr}-$F{zeile},0)",
        "ergebnis": f"=($D{zeile}-$E{zeile}-$F{zeile})*$C{zeile}",
    }
