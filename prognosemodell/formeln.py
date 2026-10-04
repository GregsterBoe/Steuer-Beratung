"""Formel-Bausteine.

Formeln werden in englischer Excel-Syntax geschrieben (so legt openpyxl sie ab);
Excel zeigt sie in der deutschen Oberfläche automatisch als WENN, ZÄHLENWENN usw.
"""

from openpyxl.utils import get_column_letter

from .modelle import (AUSWERTUNG_SPALTEN, LIQUIDITAET_SPALTEN, NEU_FELDER, NEU_SPALTEN,
                      OBJEKT_FELDER, RUECKLAGE_6B_GEBILDET, RUECKLAGE_JAHR_SPALTEN,
                      RUECKLAGE_SPALTEN, SZENARIO_B, VERKAUF_FELDER, VERKAUF_SPALTEN,
                      VERGLEICH_KENNZAHLEN)

RUECKLAGE_SZENARIO_B = "Szenario B: sofort versteuert"
NEU_SZENARIO_B = "entfällt in Szenario B"


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
    """Formeln einer Prognosezeile (Spalten A–H laut Projektplan Abschnitt 8, I–M Abschnitt 14).

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
        # Verkauf zum Jahresende: im Verkaufsjahr aktiv, am Jahresende aber nicht mehr im Bestand
        "bestand": f'=IF(AND({status}="OK",OR({verkauft}=0,$B{zeile}<{vk_jahr})),1,0)',
        "bw_gub": f"={_obj('ak_gub', objekt_zeile)}*$I{zeile}",
        "bw_gesamt": f"=($G{zeile}+$J{zeile})*$I{zeile}",
        # ohne Verkehrswert im Objektblatt gilt der Buchwert, stille Reserve also null
        "verkehrswert": f'=IF({_obj("verkehrswert", objekt_zeile)}="",$K{zeile},'
                        f'{_obj("verkehrswert", objekt_zeile)}'
                        f"*(1+par_Wertsteig)^($B{zeile}-par_Basisjahr))*$I{zeile}",
        "stille_reserven": f"=$L{zeile}-$K{zeile}",
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
    das aktiv-Flag der Prognose frei von Zirkelbezügen. Aus demselben Grund lesen
    die Formeln nur den Bestandsteil der Prognose (prgb_*): Die Zeilen der
    Neuobjekte hängen über die Rücklage am Veräußerungsgewinn.
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
            f'{sp["faktor"]}*SUMIFS(prgb_Miete,prgb_ID,{id_},prgb_Jahr,{sp["jahr"]}))'),
        "bw_geb": nur_ok(f'SUMIFS(prgb_Buchwert,prgb_ID,{id_},prgb_Jahr,{sp["jahr"]})'),
        "bw_gesamt": nur_ok(f'{sp["bw_geb"]}+{ak_gub}'),
        "erloes_geb": nur_ok(
            f'IF({methode}="Buchwert",'
            f'IF({sp["bw_gesamt"]}=0,0,{netto}*{sp["bw_geb"]}/{sp["bw_gesamt"]}),'
            f'{netto}*{quote})'),
        "gewinn_geb": nur_ok(f'{sp["erloes_geb"]}-{sp["bw_geb"]}'),
        "gewinn_gub": nur_ok(f'{netto}-{sp["erloes_geb"]}-{ak_gub}'),
        "gewinn": nur_ok(f'{sp["gewinn_geb"]}+{sp["gewinn_gub"]}'),
    }


def ruecklage_spalten() -> dict:
    """Spaltenbuchstaben des Rücklagenblatts: erst je Verkauf, nach einer Leerspalte je Jahr."""
    sp = {s.key: get_column_letter(i) for i, s in enumerate(RUECKLAGE_SPALTEN, start=1)}
    start = len(RUECKLAGE_SPALTEN) + 2
    sp.update({f"j_{s.key}": get_column_letter(i)
               for i, s in enumerate(RUECKLAGE_JAHR_SPALTEN, start=start)})
    return sp


def ruecklage_zeile(zeile: int) -> dict:
    """Rücklagenspiegel je Verkauf (Projektplan Abschnitt 12).

    Zeile n im Blatt Rücklagen gehört zu Zeile n im Blatt Verkäufe. Gerechnet wird
    nur für Verkäufe mit Status OK; die Rücklage wird je Wirtschaftsgut gebildet,
    ein Verlust bei Gebäude oder G+B bleibt sofort wirksam.
    """
    vk = {k: f"'Verkäufe'!${v}${zeile}" for k, v in verkauf_spalten().items()}
    sp = {k: f"${v}{zeile}" for k, v in ruecklage_spalten().items()}
    gebildet = f'{sp["status"]}="{RUECKLAGE_6B_GEBILDET}"'

    def nur_vk(formel: str) -> str:
        return f'=IF({sp["id"]}="","",{formel})'

    def nur_rl(formel: str) -> str:
        return f'=IF({sp["id"]}="","",IF({gebildet},{formel},0))'

    return {
        "id": f'=IF({vk["status"]}="OK",{vk["objekt_id"]},"")',
        "jahr": nur_vk(vk["jahr"]),
        "kaufjahr": nur_vk(f'INDEX(obj_Kaufjahr,MATCH({sp["id"]},obj_ID,0))'),
        "vorbesitz": nur_vk(f'{sp["jahr"]}-{sp["kaufjahr"]}'),
        # Szenario B versteuert jeden Gewinn sofort, unabhängig von der 6b-Angabe
        "status": nur_vk(
            f'IF(par_Szenario="{SZENARIO_B}","{RUECKLAGE_SZENARIO_B}",'
            f'IF({vk["nutzung_6b"]}<>"ja","6b nicht gewählt",'
            f'IF({sp["vorbesitz"]}<par_6bVorbesitz,"Vorbesitzzeit zu kurz",'
            f'IF(MAX({vk["gewinn_geb"]},0)+MAX({vk["gewinn_gub"]},0)=0,"kein Gewinn",'
            f'"{RUECKLAGE_6B_GEBILDET}"))))'),
        "ruecklage_id": f'=IF({gebildet},"R-"&{sp["id"]},"")',
        "gewinn": nur_vk(vk["gewinn"]),
        "betrag_geb": nur_rl(f'MAX({vk["gewinn_geb"]},0)'),
        "betrag_gub": nur_rl(f'MAX({vk["gewinn_gub"]},0)'),
        "ruecklage": nur_vk(f'{sp["betrag_geb"]}+{sp["betrag_gub"]}'),
        "steuerpflichtig": nur_vk(f'{sp["gewinn"]}-{sp["ruecklage"]}'),
        "fristjahr": f'=IF({gebildet},{sp["jahr"]}+par_6bFrist,"")',
        # Summe über alle Neuobjekte, die diese Rücklage als Quelle nennen
        "uebertrag_geb": nur_rl(f'SUMIFS(neu_UebGeb,neu_Quelle,{sp["ruecklage_id"]})'),
        "uebertrag_gub": nur_rl(f'SUMIFS(neu_UebGuBGuB,neu_Quelle,{sp["ruecklage_id"]})'
                                f'+SUMIFS(neu_UebGuBGeb,neu_Quelle,{sp["ruecklage_id"]})'),
        "rest": nur_vk(f'{sp["ruecklage"]}-{sp["uebertrag_geb"]}-{sp["uebertrag_gub"]}'),
        # 6 % je vollem Jahr zwischen Bildung (Ende Verkaufsjahr) und Auflösung (Ende Fristjahr)
        "zuschlag": nur_rl(f'{sp["rest"]}*par_6bZuschlag*({sp["fristjahr"]}-{sp["jahr"]})'),
    }


def ruecklage_jahr_zeile(zeile: int, jahr_index: int) -> dict:
    """Jahresspiegel aller Rücklagen; Stand = Vorjahr + Bildung − Übertrag − Auflösung."""
    sp = {k: f"${v}{zeile}" for k, v in ruecklage_spalten().items()}
    jahr = sp["j_jahr"]
    vorjahr = "0" if jahr_index == 0 else f'${ruecklage_spalten()["j_stand"]}{zeile - 1}'
    return {
        "jahr": f"=par_Startjahr+{jahr_index}",
        "gewinn": f"=SUMIFS(rl_Gewinn,rl_Jahr,{jahr})",
        "bildung": f"=SUMIFS(rl_Ruecklage,rl_Jahr,{jahr})",
        # übertragen wird im Kaufjahr des Neuobjekts
        "uebertrag": f"=SUMIFS(neu_Uebertrag,neu_Kaufjahr,{jahr})",
        "aufloesung": f"=SUMIFS(rl_Rest,rl_Fristjahr,{jahr})",
        "zuschlag": f"=SUMIFS(rl_Zuschlag,rl_Fristjahr,{jahr})",
        "stand": f'={vorjahr}+{sp["j_bildung"]}-{sp["j_uebertrag"]}-{sp["j_aufloesung"]}',
        "steuerpflichtig": f'=SUMIFS(rl_Steuerpflichtig,rl_Jahr,{jahr})'
                           f'+{sp["j_aufloesung"]}+{sp["j_zuschlag"]}',
        "steuer": f'={sp["j_steuerpflichtig"]}*par_Steuersatz',
    }


def neu_spalten() -> dict:
    """Spaltenbuchstaben des Blatts Neuobjekte: Eingaben, Status, dann Formeln."""
    sp = {f.key: get_column_letter(i) for i, f in enumerate(NEU_FELDER, start=1)}
    sp["status"] = get_column_letter(len(NEU_FELDER) + 1)
    for i, s in enumerate(NEU_SPALTEN, start=len(NEU_FELDER) + 2):
        sp[s.key] = get_column_letter(i)
    return sp


def neu_zeile(zeile: int) -> dict:
    """Status, Anschaffungskosten und Übertrag einer Zeile im Blatt Neuobjekte (Abschnitt 13).

    Übertragbarkeit nach § 6b: Die Gebäude-Rücklage geht nur auf das Gebäude, die
    G+B-Rücklage zuerst auf G+B (mindert keine AfA), der Rest auf das Gebäude.
    Übertragen wird so viel wie möglich. Nennen mehrere Neuobjekte dieselbe Rücklage,
    bedienen die Zeilen darüber zuerst; deshalb laufen die Summen bis zur Vorzeile.
    """
    sp = {k: f"${v}{zeile}" for k, v in neu_spalten().items()}
    id_, quelle, kaufjahr = sp["neu_id"], sp["quelle"], sp["kaufjahr"]
    pflicht = [sp[f.key] for f in NEU_FELDER if f.pflicht]
    treffer = f"MATCH({quelle},rl_RuecklageID,0)"
    bisher = {k: f"${neu_spalten()[k]}$1:${neu_spalten()[k]}{zeile - 1}"
              for k in ("quelle", "ueb_geb", "ueb_gub_gub", "ueb_gub_geb")}

    def vorher(key: str) -> str:
        return f'SUMIFS({bisher[key]},{bisher["quelle"]},{quelle})'

    frei_geb = f"INDEX(rl_BetragGeb,{treffer})-{vorher('ueb_geb')}"
    frei_gub = f"INDEX(rl_BetragGuB,{treffer})-{vorher('ueb_gub_gub')}-{vorher('ueb_gub_geb')}"

    def nur_ok(formel: str) -> str:
        return f'=IF({sp["status"]}<>"OK","",{formel})'

    def mit_quelle(formel: str) -> str:
        return nur_ok(f'IF({quelle}="",0,{formel})')

    return {
        "status": (
            f'=IF({id_}="","",'
            f'IF(OR(COUNTIF(neu_ID,{id_})>1,COUNTIF(obj_ID,{id_})>0),"NeuID doppelt",'
            f'IF(COUNTA({",".join(pflicht)})<{len(pflicht)},"Pflichtfeld fehlt",'
            f'IF(OR({kaufjahr}<par_Startjahr,{kaufjahr}>par_Endjahr),'
            f'"Kaufjahr außerhalb Prognose",'
            f'IF({quelle}="","OK",'
            f'IF(par_Szenario="{SZENARIO_B}","{NEU_SZENARIO_B}",'
            f'IF(COUNTIF(rl_RuecklageID,{quelle})=0,"Rücklage unbekannt",'
            f'IF({kaufjahr}<INDEX(rl_Jahr,{treffer}),"Kauf vor Verkauf",'
            f'IF({kaufjahr}>INDEX(rl_Fristjahr,{treffer}),"Kauf nach Fristjahr",'
            f'"OK")))))))))'
        ),
        "ak_gesamt": nur_ok(f'{sp["kaufpreis"]}+IF({sp["nebenkosten"]}="",'
                            f'{sp["kaufpreis"]}*par_GrESt,{sp["nebenkosten"]})'),
        "ak_geb": nur_ok(f'{sp["ak_gesamt"]}*(1-{sp["anteil_gub"]})'),
        "ak_gub": nur_ok(f'{sp["ak_gesamt"]}*{sp["anteil_gub"]}'),
        "ueb_geb": mit_quelle(f'MIN({frei_geb},{sp["ak_geb"]})'),
        "ueb_gub_gub": mit_quelle(f'MIN({frei_gub},{sp["ak_gub"]})'),
        "ueb_gub_geb": mit_quelle(
            f'MIN({frei_gub}-{sp["ueb_gub_gub"]},{sp["ak_geb"]}-{sp["ueb_geb"]})'),
        "uebertrag": nur_ok(f'{sp["ueb_geb"]}+{sp["ueb_gub_gub"]}+{sp["ueb_gub_geb"]}'),
        "afa_basis": nur_ok(f'{sp["ak_geb"]}-{sp["ueb_geb"]}-{sp["ueb_gub_geb"]}'),
        "bw_gub": nur_ok(f'{sp["ak_gub"]}-{sp["ueb_gub_gub"]}'),
    }


def prognose_neu_zeile(zeile: int, neu_zeile_nr: int, jahr_index: int) -> dict:
    """Prognosezeile eines Neuobjekts, gleiche Spalten wie beim Bestand.

    Kauf zum Jahresende: Im Kaufjahr steht nur der Buchwert (= AfA-Basis), Miete,
    Erhaltung und AfA laufen ab dem Folgejahr. Die Miete im ersten vollen Jahr ist
    Kaufpreis × Mietrendite, danach steigt sie mit der Mietsteigerung.
    """
    sp = {k: f"Neuobjekte!${v}${neu_zeile_nr}" for k, v in neu_spalten().items()}
    kaufjahr, jahr = sp["kaufjahr"], f"$B{zeile}"
    ok = f'{sp["status"]}="OK"'
    bw_vorjahr = "0" if jahr_index == 0 else f"$G{zeile - 1}"
    jahre_seit_kauf = f"({jahr}-{kaufjahr}-1)"
    return {
        "id": f'=IF({sp["neu_id"]}="","",{sp["neu_id"]})',
        "jahr": f"=par_Startjahr+{jahr_index}",
        "aktiv": f"=IF({ok},IF({jahr}>{kaufjahr},1,0),0)",
        "miete": f'=IF($C{zeile}=1,{sp["kaufpreis"]}*{sp["mietrendite"]}'
                 f"*(1+par_Mietsteig)^{jahre_seit_kauf},0)",
        "erhaltung": f'=IF($C{zeile}=1,{sp["kaufpreis"]}*{sp["erhaltungsquote"]}'
                     f"*(1+par_Erhaltsteig)^{jahre_seit_kauf},0)",
        "afa": f'=IF($C{zeile}=1,MIN({sp["afa_basis"]}*{sp["afa_satz"]},{bw_vorjahr}),0)',
        "buchwert": f'=IF({ok},IF({jahr}<{kaufjahr},0,IF({jahr}={kaufjahr},{sp["afa_basis"]},'
                    f"MAX({bw_vorjahr}-$F{zeile},0))),0)",
        "ergebnis": f"=($D{zeile}-$E{zeile}-$F{zeile})*$C{zeile}",
        # ab Ende des Kaufjahrs im Bestand; Verkehrswert = Kaufpreis, steigt ab dem Folgejahr
        "bestand": f"=IF({ok},IF({jahr}>={kaufjahr},1,0),0)",
        "bw_gub": f'=IF($I{zeile}=1,{sp["bw_gub"]},0)',
        "bw_gesamt": f"=($G{zeile}+$J{zeile})*$I{zeile}",
        "verkehrswert": f'=IF($I{zeile}=1,{sp["kaufpreis"]}'
                        f"*(1+par_Wertsteig)^({jahr}-{kaufjahr}),0)",
        "stille_reserven": f"=$L{zeile}-$K{zeile}",
    }


def jahres_spalten(spalten) -> dict:
    """Spaltenbuchstaben eines Jahresblatts (Liquidität, Auswertung)."""
    return {s.key: get_column_letter(i) for i, s in enumerate(spalten, start=1)}


def liquiditaet_zeile(zeile: int, jahr_index: int) -> dict:
    """Geldfluss je Jahr über alle Objekte, vor Finanzierung (Projektplan Abschnitt 14).

    Zahlungswirksam sind Miete und Erhaltung, nicht die AfA. Der Kauf eines
    Neuobjekts fließt vor Finanzierung voll aus Eigenmitteln ab.
    """
    sp = {k: f"${v}{zeile}" for k, v in jahres_spalten(LIQUIDITAET_SPALTEN).items()}
    jahr = sp["jahr"]

    def prg(name: str) -> str:
        return f"=SUMIFS({name},prg_Jahr,{jahr})"

    def vk(name: str) -> str:
        return f'SUMIFS({name},vk_Jahr,{jahr},vk_Status,"OK")'

    vorjahr = "0" if jahr_index == 0 else \
        f'${jahres_spalten(LIQUIDITAET_SPALTEN)["mittelzufluss_kum"]}{zeile - 1}'
    return {
        "jahr": f"=par_Startjahr+{jahr_index}",
        "miete": prg("prg_Miete"),
        "erhaltung": prg("prg_Erhaltung"),
        "afa": prg("prg_AfA"),
        "ergebnis": prg("prg_Ergebnis"),
        "ueberschuss": f'={sp["miete"]}-{sp["erhaltung"]}',
        "erloes": f'={vk("vk_PreisAngesetzt")}-{vk("vk_Kosten")}',
        "bw_rueckfluss": f'={vk("vk_BuchwertGesamt")}',
        "gewinn": f'={vk("vk_Gewinn")}',
        # ohne Verlustvortrag: ein Verlust mindert die Steuer im selben Jahr
        "steuer_laufend": f'={sp["ergebnis"]}*par_Steuersatz',
        "steuer_verkauf": f"=SUMIFS(rlj_Steuer,rlj_Jahr,{jahr})",
        "steuer": f'={sp["steuer_laufend"]}+{sp["steuer_verkauf"]}',
        "reinvest": f'=SUMIFS(neu_AKGesamt,neu_Kaufjahr,{jahr},neu_Status,"OK")',
        "mittelzufluss": f'={sp["ueberschuss"]}+{sp["erloes"]}-{sp["steuer"]}-{sp["reinvest"]}',
        "mittelzufluss_kum": f'={vorjahr}+{sp["mittelzufluss"]}',
    }


def auswertung_zeile(zeile: int, jahr_index: int) -> dict:
    """Kennzahlen je Jahr: Gesamt-GuV, Steuer, Buch- und Verkehrswert, stille Reserven."""
    bst = jahres_spalten(AUSWERTUNG_SPALTEN)
    sp = {k: f"${v}{zeile}" for k, v in bst.items()}
    jahr = sp["jahr"]
    vorjahr = "0" if jahr_index == 0 else f'${bst["steuer_kum"]}{zeile - 1}'
    anlage_vorjahr = "0" if jahr_index == 0 else f'${bst["anlage"]}{zeile - 1}'
    return {
        "jahr": f"=par_Startjahr+{jahr_index}",
        "ergebnis": f"=SUMIFS(liq_Ergebnis,liq_Jahr,{jahr})",
        "steuerpflichtig_vk": f"=SUMIFS(rlj_Steuerpflichtig,rlj_Jahr,{jahr})",
        "guv": f'={sp["ergebnis"]}+{sp["steuerpflichtig_vk"]}',
        "steuer": f"=SUMIFS(liq_Steuer,liq_Jahr,{jahr})",
        "nach_steuer": f'={sp["guv"]}-{sp["steuer"]}',
        "steuer_kum": f'={vorjahr}+{sp["steuer"]}',
        "buchwert": f"=SUMIFS(prg_BuchwertGesamt,prg_Jahr,{jahr})",
        "verkehrswert": f"=SUMIFS(prg_Verkehrswert,prg_Jahr,{jahr})",
        "stille_reserven": f'={sp["verkehrswert"]}-{sp["buchwert"]}',
        "ruecklage": f"=SUMIFS(rlj_Stand,rlj_Jahr,{jahr})",
        "mittel_kum": f"=SUMIFS(liq_MittelzuflussKum,liq_Jahr,{jahr})",
        # Alternativanlage: Bestand Vorjahresende verzinst, Mittelzufluss zum Jahresende
        "zins": f"={anlage_vorjahr}*par_Alternativrendite",
        "steuer_zins": f'={sp["zins"]}*par_Steuersatz',
        "anlage": f'={anlage_vorjahr}+SUMIFS(liq_Mittelzufluss,liq_Jahr,{jahr})'
                  f'+{sp["zins"]}-{sp["steuer_zins"]}',
        "latente_steuer": f'=({sp["stille_reserven"]}+{sp["ruecklage"]})*par_Steuersatz',
        "vermoegen": f'={sp["verkehrswert"]}+{sp["anlage"]}-{sp["latente_steuer"]}',
    }


def vergleich_zeilen(erste: int) -> dict:
    """Zeilennummer je Kennzahl im Blatt Vergleich."""
    return {k.key: zeile for zeile, k in enumerate(VERGLEICH_KENNZAHLEN, start=erste)}


def vergleich_aktuell(erste: int, spalte: str = "B") -> dict:
    """Kennzahlen des aktiven Szenarios: Bestände am Ende des letzten Jahres, sonst Summen."""
    z = {k: f"{spalte}{n}" for k, n in vergleich_zeilen(erste).items()}

    def ende(name: str) -> str:
        return f"=SUMIFS({name},aw_Jahr,par_Endjahr)"

    return {
        "verkehrswert": ende("aw_Verkehrswert"),
        "anlage": ende("aw_Anlage"),
        "latente_steuer": ende("aw_LatenteSteuer"),
        "vermoegen": f'={z["verkehrswert"]}+{z["anlage"]}-{z["latente_steuer"]}',
        "buchwert": ende("aw_Buchwert"),
        "stille_reserven": ende("aw_StilleReserven"),
        "ruecklage": ende("aw_Ruecklage"),
        "miete": "=SUM(liq_Miete)",
        "afa": "=SUM(liq_AfA)",
        "ergebnis": "=SUM(liq_Ergebnis)",
        "steuer": "=SUM(aw_Steuer)",
        "zins": "=SUM(aw_Zins)",
        "steuer_zins": "=SUM(aw_SteuerZins)",
        "steuer_gesamt": f'={z["steuer"]}+{z["steuer_zins"]}',
        "reinvest": "=SUM(liq_Reinvest)",
    }
