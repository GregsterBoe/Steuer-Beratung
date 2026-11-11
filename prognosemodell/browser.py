"""Browser-Fassung: eine HTML-Datei mit Rechenwerk, Oberfläche und Daten, ohne Setup.

Aufruf:
    python -m prognosemodell.browser                       # ausgabe/Prognosemodell.html mit Beispiel
    python -m prognosemodell.browser --ohne-testdaten      # leer
    python -m prognosemodell.browser --mappe Prognose.xlsx # Eingaben einer Excel-Mappe übernehmen

Die Seite rechnet mit browser/rechenwerk.js dieselben Formeln wie die Mappe
(pruefung/abgleich.py prüft das). Bezeichnungen, Erläuterungen, Parameter und Prüfungen
kommen aus modelle.py, damit Mappe und Seite dieselben Begriffe zeigen.
"""

import argparse
import dataclasses
import datetime
import json
import re
import sys
from pathlib import Path

from .modelle import (ANLAGE_FELDER, ANLAGE_SPALTEN, DARLEHEN_SPALTEN, FMT_EURO, FMT_FAKTOR,
                      FMT_JAHR, FMT_PROZENT, FMT_RENDITE, FMT_TEXT, KST_VERKNUEPFT, NEU_FELDER,
                      NEU_SPALTEN, NEUKAUF_POSITIONEN, OBJEKT_FELDER, PARAMETER, PROGNOSE_SPALTEN,
                      PRUEFUNGEN, RUECKLAGE_JAHR_SPALTEN, RUECKLAGE_SPALTEN, SZENARIEN,
                      VERGLEICH_KENNZAHLEN, VERKAUF_FELDER, VERKAUF_SPALTEN, Neuobjekt, Objekt,
                      Verkauf, aus_spalten, liq_spalten)
from .testdaten import testobjekt

ORDNER = Path(__file__).resolve().parent.parent / "browser"
STANDARD_AUSGABE = Path("ausgabe") / "Prognosemodell.html"


# ---- Metadaten aus modelle.py ----

def _format(fmt: str) -> str:
    return {FMT_EURO: "euro", FMT_PROZENT: "prozent", FMT_JAHR: "jahr", FMT_RENDITE: "prozent1",
            FMT_FAKTOR: "faktor", FMT_TEXT: "text", '#,##0.00 "€"': "euro2", "0%": "prozent0",
            "0.00": "dez2", "0.0": "dez1", "DD.MM.YYYY": "datum"}.get(fmt, "zahl")


def _feld(f) -> dict:
    return {"key": f.key, "ueberschrift": f.ueberschrift, "name": f.name, "format": _format(f.format),
            "pflicht": f.pflicht, "auswahl": list(f.auswahl) if f.auswahl else None,
            "annahme": f.annahme, "hinweis": f.hinweis, "berechnet": f.berechnet,
            "minimum": f.minimum, "maximum": f.maximum}


def _spalte(s) -> dict:
    return {"key": s.key, "ueberschrift": s.ueberschrift, "name": s.name, "format": _format(s.format)}


def metadaten() -> dict:
    parameter = []
    for p in PARAMETER:
        berechnet = isinstance(p.wert, str) and p.wert.startswith("=")
        parameter.append({
            "name": p.name[4:], "bezeichnung": p.bezeichnung, "format": _format(p.format),
            "wert": None if berechnet else p.wert, "berechnet": berechnet,
            "erlaeuterung": p.erlaeuterung, "auswahl": list(p.auswahl) if p.auswahl else None,
            "abschnitt": p.abschnitt})
    neu_spalten = [_spalte(s) for s in NEU_SPALTEN]
    return {
        "parameter": parameter,
        "felder": {"objekte": [_feld(f) for f in OBJEKT_FELDER],
                   "verkaeufe": [_feld(f) for f in VERKAUF_FELDER],
                   "neuobjekte": [_feld(f) for f in NEU_FELDER],
                   "anlagen": [_feld(f) for f in ANLAGE_FELDER]},
        "spalten": {"prognose": [_spalte(s) for s in PROGNOSE_SPALTEN],
                    "verkauf": [_spalte(s) for s in VERKAUF_SPALTEN],
                    "ruecklage": [_spalte(s) for s in RUECKLAGE_SPALTEN],
                    "spiegel": [_spalte(s) for s in RUECKLAGE_JAHR_SPALTEN],
                    "neu": neu_spalten,
                    "darlehen": [_spalte(s) for s in DARLEHEN_SPALTEN],
                    "anlage": [_spalte(s) for s in ANLAGE_SPALTEN],
                    "liq": [_spalte(s) for s in liq_spalten(SZENARIEN[0])],
                    "aus": [_spalte(s) for s in aus_spalten(SZENARIEN[0])]},
        "szenarien": [{"key": s.key, "titel": s.titel, "kurz": s.kurz, "liq": s.liq, "aus": s.aus}
                      for s in SZENARIEN],
        "vergleich": [{"kennzahl": k[0], "erlaeuterung": k[4]} for k in VERGLEICH_KENNZAHLEN],
        "pruefungen": [{"key": p.key, "bezeichnung": p.bezeichnung, "art": p.art, "wo": p.wo}
                       for p in PRUEFUNGEN],
        "neukauf_positionen": [{"key": k, "bezeichnung": b, "bwa": nr}
                               for k, b, nr, _ in NEUKAUF_POSITIONEN],
    }


def parameter_vorgaben() -> dict:
    """Parameter ohne „par_“; berechnete (Formeln) rechnet das Rechenwerk selbst."""
    return {p.name[4:]: (None if isinstance(p.wert, str) and p.wert.startswith("=") else p.wert)
            for p in PARAMETER}


# ---- Eingaben ----

def _json_wert(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.isoformat()[:10]
    return v


def als_dict(obj) -> dict:
    return {k: _json_wert(v) for k, v in dataclasses.asdict(obj).items()}


def beispiel() -> dict:
    """Ein Bestandsobjekt mit Darlehen, ein zweites, ein Verkauf mit § 6b und ein Neuobjekt
    mit Kredit aus dessen Rücklage."""
    obj = testobjekt()
    objekte = [dataclasses.replace(obj, restschuld=500_000, zinssatz=0.04, rate=40_000),
               dataclasses.replace(obj, objekt_id="OBJ-002", name="Zweites Objekt",
                                   weitere_einnahmen=1_000, weitere_ausgaben=2_000)]
    verkaeufe = [Verkauf("OBJ-001", 2029, preis=1_400_000, kosten=30_000, nutzung_6b="ja")]
    neuobjekte = [Neuobjekt("NEU-001", 2030, name="Neubau Mehrfamilienhaus", kaufpreis=1_500_000,
                            anteil_gub=0.2, nebenkosten=105_000, afa_satz=0.03,
                            mietrendite=0.045, erhaltungsquote=0.005, quelle="RL-OBJ-001",
                            erloes_fin="ja", fin_art="Kredit", kredit_quote=0.6,
                            kredit_zins=0.04, tilgung=0.02)]
    return leere_eingabe(objekte=[als_dict(o) for o in objekte],
                         verkaeufe=[als_dict(v) for v in verkaeufe],
                         neuobjekte=[als_dict(n) for n in neuobjekte])


def leere_eingabe(**daten) -> dict:
    return {"parameter": parameter_vorgaben(), "objekte": [], "verkaeufe": [], "neuobjekte": [],
            "anlagen": [], "afa_plan": {}, "neukauf": [], "schnellcheck": False, **daten}


# ---- Eingaben aus einer Excel-Mappe ----

def bereich(wb, name: str) -> list:
    """Alle Zellen eines benannten Bereichs als Liste von Zeilen (Zellen, nicht Werte)."""
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    zellen = wb[blatt][ref.replace("$", "")]
    if not isinstance(zellen, tuple):
        return [[zellen]]
    if not isinstance(zellen[0], tuple):
        zellen = (zellen,)
    return [list(z) for z in zellen]


def werte(wb, name: str) -> list:
    return [[c.value for c in z] for z in bereich(wb, name)]


def spalte(wb, name: str) -> list:
    return [z[0] for z in werte(wb, name)]


def _ist_formel(zelle) -> bool:
    return isinstance(zelle.value, str) and zelle.value.startswith("=")


def _eingetippt(zelle):
    """Wert einer Eingabezelle ohne Formel; leer heißt bewusst leer (""), nicht Annahme."""
    return "" if zelle.value is None else _json_wert(zelle.value)


def ist_gerechnet(wv) -> bool:
    """Hat die Mappe gespeicherte Formelergebnisse? Eine frisch erzeugte Mappe hat keine:
    Python schreibt nur die Formeln, erst Excel rechnet sie beim Öffnen und Speichern."""
    return spalte(wv, "prg_Jahr")[0] is not None


def startjahr(wf, wv) -> int:
    """Erstes Prognosejahr; ohne gespeicherte Formelergebnisse aus den Parametern."""
    gerechnet = spalte(wv, "prg_Jahr")[0]
    if gerechnet is not None:
        return int(gerechnet)
    vorgaben = {p.name: p.wert for p in PARAMETER}

    def eingetippt(name):
        wert = bereich(wf, name)[0][0].value if name in wf.defined_names else vorgaben[name]
        return wert if isinstance(wert, (int, float)) else None

    start = eingetippt("par_Startjahr")
    if start is None:
        start = (eingetippt("par_Basisjahr") or vorgaben["par_Basisjahr"]) + 1
    return int(start)


def afa_plan(wb, start: int) -> dict:
    """AfA-Plan je ObjektID aus einer gerechneten Mappe (Werte)."""
    plan = {}
    for id_, reihe in zip(spalte(wb, "afp_ID"), werte(wb, "afp_Jahre")):
        jahre = {start + i: w for i, w in enumerate(reihe) if isinstance(w, (int, float))}
        if id_ and jahre:
            plan[str(id_)] = jahre
    return plan


def neukauf(wb, start: int) -> list:
    """Neukauf-Kostenstellen mit ihren eingetragenen Jahreswerten (Werte)."""
    zeilen = werte(wb, "nk_Eingabe")
    basis = start - 1
    je = len(NEUKAUF_POSITIONEN)
    liste = []
    for kopf in range(0, len(zeilen), je):
        kst = zeilen[kopf][0]
        if not kst:
            continue
        jahre = {}
        for k, (key, *_) in enumerate(NEUKAUF_POSITIONEN):
            reihe = zeilen[kopf + k][4:]
            jahre[key] = {basis + i: w for i, w in enumerate(reihe) if isinstance(w, (int, float))}
        liste.append({"kst": kst, "name": zeilen[kopf][1], "werte": jahre})
    return liste


def eingabe_aus_mappe(pfad: Path) -> dict:
    """Eingaben einer gespeicherten Mappe: eingetippte Werte bleiben Eingaben, Annahmeformeln
    (blau) bleiben leer, damit die Seite sie selbst rechnet. Verknüpfungen auf die
    Kostenstellenblätter kommen als Werte (in Excel gespeichert, sonst fehlen sie)."""
    from openpyxl import load_workbook
    wf = load_workbook(pfad)
    wv = load_workbook(pfad, data_only=True)

    def zeilen(felder, erste_key):
        formeln = {f.key: [z[0] for z in bereich(wf, f.name)] for f in felder}
        gerechnet = {f.key: spalte(wv, f.name) for f in felder}
        n = len(formeln[erste_key])
        return formeln, gerechnet, n

    param = parameter_vorgaben()
    for p in PARAMETER:
        if p.name not in wf.defined_names:
            continue
        zelle = bereich(wf, p.name)[0][0]
        if not _ist_formel(zelle) and zelle.value is not None:
            param[p.name[4:]] = _json_wert(zelle.value)
        elif not _ist_formel(zelle):
            param[p.name[4:]] = None

    def bis_letzte(werte_id):
        """Zeilen bis zur letzten gefüllten; leere dazwischen bleiben stehen, denn die
        Zeilennummer verbindet Verkauf und Reinvestition."""
        gefuellt = [j for j, w in enumerate(werte_id) if w not in (None, "")]
        return range(gefuellt[-1] + 1) if gefuellt else range(0)

    objekte = []
    formeln, gerechnet, n = zeilen(OBJEKT_FELDER, "objekt_id")
    for j in bis_letzte(gerechnet["objekt_id"]):
        if gerechnet["objekt_id"][j] in (None, ""):
            objekte.append({"objekt_id": None})
            continue
        d, herkunft = {}, {}
        for f in OBJEKT_FELDER:
            zelle = formeln[f.key][j]
            if not _ist_formel(zelle):
                d[f.key] = _eingetippt(zelle)
            elif f.key in KST_VERKNUEPFT and "!" in zelle.value:
                wert = gerechnet[f.key][j]
                d[f.key] = "" if wert is None else wert
                herkunft[f.key] = "kostenstelle"
            else:
                d[f.key] = None          # Annahme rechnet die Seite
        if herkunft:
            d["herkunft"] = herkunft
        objekte.append(d)

    verkaeufe = []
    felder = [f for f in VERKAUF_FELDER if f.key != "bezeichnung"]
    formeln, gerechnet, n = zeilen(felder, "objekt_id")
    for j in bis_letzte(gerechnet["objekt_id"]):
        verkaeufe.append({f.key: (None if _ist_formel(formeln[f.key][j])
                                  else _eingetippt(formeln[f.key][j])) for f in felder})

    neuobjekte = []
    formeln, gerechnet, n = zeilen(NEU_FELDER, "neu_id")
    for r in range(n):
        if _ist_formel(formeln["neu_id"][r]):
            neuobjekte.append(None)      # Annahmezeile (reinvestieren = ja) rechnet die Seite
            continue
        if formeln["neu_id"][r].value in (None, ""):
            neuobjekte.append(None)
            continue
        d = {}
        for f in NEU_FELDER:
            zelle = formeln[f.key][r]
            if f.key in ("ak_gub_eingabe", "ak_geb_eingabe") and _ist_formel(zelle):
                d[f.key] = None
            elif _ist_formel(zelle):
                wert = gerechnet[f.key][r]
                d[f.key] = None if wert in (None, "") else _json_wert(wert)
            else:
                d[f.key] = _eingetippt(zelle)
        neuobjekte.append(d)
    while neuobjekte and neuobjekte[-1] is None:
        neuobjekte.pop()

    anlagen = []
    formeln, gerechnet, n = zeilen(ANLAGE_FELDER, "nr")
    for j in bis_letzte(gerechnet["nr"]):
        anlagen.append({f.key: (None if _ist_formel(formeln[f.key][j])
                                else _eingetippt(formeln[f.key][j])) for f in ANLAGE_FELDER})

    # Schnellcheck: § 6b und reinvestieren stehen als Annahmeformel in den Verkaufszeilen
    schnellcheck = any(_ist_formel(z[0]) for z in bereich(wf, "vk_6b"))
    if not ist_gerechnet(wv):
        print(f"Hinweis: {Path(pfad).name} enthält keine gerechneten Werte (noch nie in Excel "
              "gespeichert). Eingetippte Werte sind übernommen; Werte, die in der Mappe aus "
              "Formeln kommen (Verknüpfungen auf Kostenstellenblätter), fehlen. Dafür die "
              "Mappe einmal in Excel öffnen, speichern und den Befehl wiederholen.",
              file=sys.stderr)
    start = startjahr(wf, wv)
    return leere_eingabe(parameter=param, objekte=objekte, verkaeufe=verkaeufe,
                         neuobjekte=neuobjekte, anlagen=anlagen, afa_plan=afa_plan(wv, start),
                         neukauf=neukauf(wv, start), quelle=Path(pfad).name,
                         schnellcheck=schnellcheck)


# ---- HTML ----

def _json_im_skript(daten) -> str:
    return json.dumps(daten, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


def erstelle_html(eingabe: dict, titel: str = "Prognosemodell V+V") -> str:
    vorlage = (ORDNER / "seite.html").read_text(encoding="utf-8")
    stil = "\n".join((ORDNER / name).read_text(encoding="utf-8")
                     for name in ("stil_werte.css", "stil.css"))
    skripte = "\n".join((ORDNER / name).read_text(encoding="utf-8")
                        for name in ("rechenwerk.js", "rechenweg.js", "oberflaeche.js"))
    daten = (f"window.PROGNOSE_META={_json_im_skript(metadaten())};\n"
             f"window.PROGNOSE_START={_json_im_skript(eingabe)};\n"
             f"window.PROGNOSE_ERZEUGT={_json_im_skript(datetime.date.today().isoformat())};")
    ersetzt = {"<!--TITEL-->": titel, "/*STIL*/": stil, "/*DATEN*/": daten,
               "/*SKRIPTE*/": skripte.replace("</script", "<\\/script")}
    return re.sub("|".join(re.escape(k) for k in ersetzt), lambda m: ersetzt[m.group(0)], vorlage)


STIL_ENDE = "/* House style: base */"


def uebernimm_stil(css: str) -> Path:
    """Schreibt browser/stil_werte.css: nur die Werte des Stils (Farben, Schriften, Rundungen,
    Skalen), ohne dessen Bausteine. Die Seite ordnet sie in stil.css ihren Rollen zu."""
    if STIL_ENDE not in css or "--paper:" not in css:
        raise SystemExit("Das ist nicht die Ausgabe von `python -m kb style ID --css`.")
    werte = css.split(STIL_ENDE)[0].rstrip() + "\n"
    ziel = ORDNER / "stil_werte.css"
    ziel.write_text("/* Erzeugt mit python -m prognosemodell.browser --stil, nicht von Hand ändern. */\n"
                    + werte, encoding="utf-8")
    return ziel


def main(argv=None) -> int:
    argumente = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    argumente.add_argument("--mappe", type=Path, help="Eingaben aus dieser Excel-Mappe übernehmen")
    argumente.add_argument("--ohne-testdaten", action="store_true", help="leere Seite")
    argumente.add_argument("--ausgabe", type=Path, default=STANDARD_AUSGABE)
    argumente.add_argument("--stil", type=Path, metavar="CSS",
                           help="Werte eines Stils der Knowledge-Base übernehmen "
                                "(Ausgabe von `python -m kb style ID --css`) und beenden")
    args = argumente.parse_args(argv)
    if args.stil:
        print(f"geschrieben: {uebernimm_stil(args.stil.read_text(encoding='utf-8'))}")
        return 0
    if args.mappe:
        eingabe = eingabe_aus_mappe(args.mappe)
    elif args.ohne_testdaten:
        eingabe = leere_eingabe()
    else:
        eingabe = beispiel()
    ziel = args.ausgabe
    if ziel.suffix.lower() != ".html":
        ziel = ziel / STANDARD_AUSGABE.name
    ziel.parent.mkdir(parents=True, exist_ok=True)
    ziel.write_text(erstelle_html(eingabe), encoding="utf-8")
    print(f"geschrieben: {ziel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
