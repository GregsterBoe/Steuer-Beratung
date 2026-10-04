"""Formel-Bausteine.

Formeln werden in englischer Excel-Syntax geschrieben (so legt openpyxl sie ab);
Excel zeigt sie in der deutschen Oberfläche automatisch als WENN, ZÄHLENWENN usw.
"""

from openpyxl.utils import get_column_letter

from .modelle import OBJEKT_FELDER, VERKAUF_FELDER, VERKAUF_SPALTEN


def spalte(key: str, felder=OBJEKT_FELDER) -> str:
    """Spaltenbuchstabe eines Eingabefelds (Standard: Objektblatt)."""
    for i, f in enumerate(felder, start=1):
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
    id_ = _obj("objekt_id", objekt_zeile)
    # Verkaufsjahr nur aus gültigen Verkäufen; ohne Verkauf bleibt das Objekt aktiv
    verkauft = f'COUNTIFS(vk_ID,{id_},vk_Status,"OK")'
    vk_jahr = f'SUMIFS(vk_Jahr,vk_ID,{id_},vk_Status,"OK")'
    bw_vorjahr = _obj("restbuchwert", objekt_zeile) if jahr_index == 0 else f"$G{zeile - 1}"
    afa_voll = f"{_obj('ak_gebaeude', objekt_zeile)}*{_obj('afa_satz', objekt_zeile)}"
    return {
        "id": f'=IF({_obj("objekt_id", objekt_zeile)}="","",{_obj("objekt_id", objekt_zeile)})',
        "jahr": f"=par_Startjahr+{jahr_index}",
        # nur Objekte mit Status OK, und nur bis einschließlich Verkaufsjahr
        "aktiv": f'=IF(AND({status}="OK",OR({verkauft}=0,$B{zeile}<={vk_jahr})),1,0)',
        "miete": f"={_obj('miete', objekt_zeile)}*(1+par_Mietsteig)^($B{zeile}-par_Basisjahr)"
                 f"*$C{zeile}",
        "erhaltung": f"={_obj('erhaltung', objekt_zeile)}"
                     f"*(1+par_Erhaltsteig)^($B{zeile}-par_Basisjahr)*$C{zeile}",
        "afa": f"=MIN({afa_voll},{bw_vorjahr})*$C{zeile}",
        # nach dem Verkauf (oder ohne gültiges Objekt) steht kein Buchwert mehr
        "buchwert": f"=MAX({bw_vorjahr}-$F{zeile},0)*$C{zeile}",
        "ergebnis": f"=($D{zeile}-$E{zeile}-$F{zeile})*$C{zeile}",
    }


def verkauf_spalten() -> dict:
    """Spaltenbuchstaben des Verkaufsblatts: Eingaben, Status, dann Formeln."""
    sp = {f.key: get_column_letter(i) for i, f in enumerate(VERKAUF_FELDER, start=1)}
    sp["status"] = get_column_letter(len(VERKAUF_FELDER) + 1)
    for i, s in enumerate(VERKAUF_SPALTEN, start=len(VERKAUF_FELDER) + 2):
        sp[s.key] = get_column_letter(i)
    return sp


def verkauf_zeile(zeile: int) -> dict:
    """Status und Formeln einer Zeile im Blatt Verkäufe (Projektplan Abschnitt 11).

    Der Status prüft nur Eingaben und Objektblatt, nicht die Prognose; so bleibt
    das aktiv-Flag der Prognose frei von Zirkelbezügen.
    """
    sp = {k: f"${v}{zeile}" for k, v in verkauf_spalten().items()}
    id_ = sp["objekt_id"]
    treffer = f"MATCH({id_},obj_ID,0)"
    ak_gub = f"INDEX(obj_AKGuB,{treffer})"
    quote = f"INDEX(obj_VKQuoteGeb,{treffer})"
    methode = f'IF({sp["aufteilung"]}="",par_Aufteilung,{sp["aufteilung"]})'
    netto = f'({sp["preis_angesetzt"]}-{sp["kosten"]})'

    def nur_ok(formel: str) -> str:
        return f'=IF({sp["status"]}<>"OK","",{formel})'

    return {
        "status": (
            f'=IF({id_}="","",'
            f'IF(COUNTIF(obj_ID,{id_})=0,"ObjektID unbekannt",'
            f'IF(INDEX(obj_Status,{treffer})<>"OK","Objekt nicht OK",'
            f'IF(COUNTIF(vk_ID,{id_})>1,"Verkauf doppelt",'
            f'IF(OR({sp["jahr"]}="",{sp["nutzung_6b"]}=""),"Pflichtfeld fehlt",'
            f'IF(COUNTA({sp["preis"]},{sp["faktor"]})<>1,"Preis oder Faktor angeben",'
            f'IF(OR({sp["jahr"]}<par_Startjahr,{sp["jahr"]}>par_Endjahr),'
            f'"Verkaufsjahr außerhalb Prognose",'
            f'IF(AND({methode}="Verkehrswert",COUNTIFS(obj_ID,{id_},obj_VKQuoteGeb,"<>")=0),'
            f'"Gebäudeanteil fehlt",'
            f'"OK"))))))))'
        ),
        # Faktor bezieht sich auf die Miete im Verkaufsjahr laut Prognose
        "preis_angesetzt": nur_ok(
            f'IF({sp["preis"]}<>"",{sp["preis"]},'
            f'{sp["faktor"]}*SUMIFS(prg_Miete,prg_ID,{id_},prg_Jahr,{sp["jahr"]}))'),
        "bw_geb": nur_ok(f'SUMIFS(prg_Buchwert,prg_ID,{id_},prg_Jahr,{sp["jahr"]})'),
        "bw_gesamt": nur_ok(f'{sp["bw_geb"]}+{ak_gub}'),
        "erloes_geb": nur_ok(
            f'IF({methode}="Buchwert",'
            f'IF({sp["bw_gesamt"]}=0,0,{netto}*{sp["bw_geb"]}/{sp["bw_gesamt"]}),'
            f'{netto}*{quote})'),
        "gewinn_geb": nur_ok(f'{sp["erloes_geb"]}-{sp["bw_geb"]}'),
        "gewinn_gub": nur_ok(f'{netto}-{sp["erloes_geb"]}-{ak_gub}'),
        "gewinn": nur_ok(f'{sp["gewinn_geb"]}+{sp["gewinn_gub"]}'),
    }
