"""Abgleich Excel gegen Browser: dieselben Fälle in der Mappe (LibreOffice) und im
Rechenwerk der HTML-Seite (browser/rechenwerk.js, per Node) rechnen und Zahl für Zahl
vergleichen.

Aufruf: python -m pruefung.abgleich [-j PROZESSE] [-v] [Teil des Fallnamens]

Die Fälle kommen aus pruefung.pruefen. Verglichen werden alle Rechenblätter: Objekte (mit
Auffülllogik), Anlagen, Prognose, Verkäufe, Rücklagen mit Spiegel, Neuobjekte, Darlehen,
Liquidität und Auswertung je Szenario, Vergleich und die Anzahl je Prüfung. Fälle mit
Makros, nachträglich getippten Zellen oder Übernahme aus einer alten Mappe lassen sich
nicht als Modell an das Rechenwerk geben und entfallen.

Aus der Mappe statt aus dem Modell kommen nur, was die Mappe aus den Kostenstellenblättern
verknüpft (Basiswerte der Objekte mit Kostenstelle, Basisjahr der Neukauf-KSt) und der
AfA-Plan: die Kostenstellenblätter gibt es im Browser nicht, dort sind es Eingaben.
"""

import argparse
import concurrent.futures
import dataclasses
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl.utils import get_column_letter

from prognosemodell import formeln
from prognosemodell.browser import eingabe_aus_mappe
from prognosemodell.bwa import OBJEKT_AUS_KST
from prognosemodell.modelle import (ANLAGE_FELDER, ANLAGE_JAHRE_NAME, ANLAGE_SPALTEN,
                                    ANLAGE_STATUS_NAME, DARLEHEN_BLOECKE, DARLEHEN_SPALTEN,
                                    NEU_FELDER, NEU_SPALTEN, NEU_STATUS_NAME, NEUKAUF_POSITIONEN,
                                    OBJEKT_ANLAGEN_SPALTEN, OBJEKT_FELDER, PARAMETER,
                                    PROGNOSE_SPALTEN, PRUEFUNGEN, QUELLEN, RUECKLAGE_JAHR_SPALTEN,
                                    RUECKLAGE_SPALTEN, STATUS_NAME, SZENARIEN, VERGLEICH_KENNZAHLEN,
                                    VERKAUF_FELDER, VERKAUF_SPALTEN, VERKAUF_STATUS_NAME, Modell,
                                    aus_spalten, liq_spalten, prognosejahre)
from pruefung import pruefen

WURZEL = Path(__file__).resolve().parent.parent
CLI = WURZEL / "browser" / "rechne_cli.js"
TOLERANZ = 0.01
# Prüfungen, die im Browser nichts zu prüfen haben (BWA-Blätter, Kostenstellenblätter,
# Ergebnisbrücke, Cashflow-Blatt)
OHNE_BROWSER = {"bwa_zuordnung", "kst_abweichung", "bruecke", "cashflow"}


# ---- Eingabe für das Rechenwerk ----

def _json_wert(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.isoformat()[:10]
    return v


def _als_dict(obj) -> dict:
    return {k: _json_wert(v) for k, v in dataclasses.asdict(obj).items()}


def parameter_vorgaben() -> dict:
    """Parameter vom Parameterblatt ohne „par_“; berechnete (Formeln) rechnet das Rechenwerk."""
    return {p.name[4:]: (None if isinstance(p.wert, str) and p.wert.startswith("=") else p.wert)
            for p in PARAMETER}


def eingabe(modell: Modell, wb) -> dict:
    """Eingabe des Rechenwerks aus dem Modell, ergänzt um die Verknüpfungen der Mappe."""
    param = parameter_vorgaben()
    param = {k: v for k, v in param.items() if v is not None or k in (
        "AusschuettungAb", "MindestKDD")}
    for name, wert in modell.parameter.items():
        param[name[4:] if name.startswith("par_") else name] = wert
    objekte = []
    leer_kst = set()
    for j, obj in enumerate(modell.objekte):
        d = _als_dict(obj)
        if obj.objekt_id and obj.objekt_id in modell.kostenstellen:
            herkunft = {}
            for key in OBJEKT_AUS_KST:
                name = next(f.name for f in OBJEKT_FELDER if f.key == key)
                wert = spalte(wb, name)[j]
                if wert is None or _bezug_leer(wb, name, j):
                    # wie LibreOffice: leer, und die Formel zählt als Annahme
                    d[key], herkunft[key] = "", "annahme"
                    leer_kst.add((j, name))
                else:
                    d[key], herkunft[key] = wert, "kostenstelle"
            d["herkunft"] = herkunft
        objekte.append(d)
    return {
        "parameter": param,
        "objekte": objekte,
        "verkaeufe": [_als_dict(v) for v in modell.verkaeufe],
        "neuobjekte": [_als_dict(n) for n in modell.neuobjekte],
        "anlagen": [_als_dict(a) for a in modell.anlagen],
        "afa_plan": afa_plan(wb),
        "neukauf": neukauf(wb),
        "schnellcheck": modell.schnellcheck,
        # nur für den Vergleich: Felder mit leerem Bezug, Excel zeigt 0
        "_leer_kst": sorted(leer_kst),
    }


def _bezug_leer(wb, name: str, j: int) -> bool:
    """Die Zelle verweist nur auf eine leere Zelle des Kostenstellenblatts. LibreOffice
    vergleicht einen solchen Bezug gleich "" (Pflichtfeld fehlt, Annahme), Excel als 0; das
    Rechenwerk bekommt dann wie LibreOffice ein leeres Feld."""
    erzeugt = getattr(wb, "erzeugt", None)
    if erzeugt is None:
        return False
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    erste = erzeugt[blatt][ref.replace("$", "").split(":")[0]]
    formel = erzeugt[blatt].cell(row=erste.row + j, column=erste.column).value
    treffer = re.fullmatch(r"='((?:[^']|'')+)'!\$?([A-Z]+)\$?(\d+)", str(formel or ""))
    if not treffer:
        return False
    ziel = treffer.group(1).replace("''", "'")
    return wb[ziel][f"{treffer.group(2)}{treffer.group(3)}"].value in (None, "")


def afa_plan(wb) -> dict:
    ids, jahre = spalte(wb, "afp_ID"), bereich(wb, "afp_Jahre")
    start = jahr0(wb)
    plan = {}
    for id_, reihe in zip(ids, jahre):
        werte = {start + i: w for i, w in enumerate(reihe) if isinstance(w, (int, float))}
        if id_ and werte:
            plan[str(id_)] = werte
    return plan


def neukauf(wb) -> list:
    zeilen = bereich(wb, "nk_Eingabe")
    basis = jahr0(wb) - 1
    je = len(NEUKAUF_POSITIONEN)
    liste = []
    for kopf in range(0, len(zeilen), je):
        kst = zeilen[kopf][0]
        if not kst:
            continue
        werte = {}
        for k, (key, *_) in enumerate(NEUKAUF_POSITIONEN):
            reihe = zeilen[kopf + k][4:]
            werte[key] = {basis + i: w for i, w in enumerate(reihe) if isinstance(w, (int, float))}
        liste.append({"kst": kst, "name": zeilen[kopf][1], "werte": werte})
    return liste


def jahr0(wb) -> int:
    return int(spalte(wb, "prg_Jahr")[0])


# ---- Werte aus der gerechneten Mappe ----

def bereich(wb, name: str) -> list:
    """Alle Zellen eines benannten Bereichs als Liste von Zeilen."""
    blatt, ref = next(iter(wb.defined_names[name].destinations))
    zellen = wb[blatt][ref.replace("$", "")]
    if not isinstance(zellen, tuple):
        return [[zellen.value]]
    if not isinstance(zellen[0], tuple):
        zellen = (zellen,)
    return [[c.value for c in z] for z in zellen]


def spalte(wb, name: str) -> list:
    return [z[0] for z in bereich(wb, name)]


# ---- Vergleich ----

def gleich(ist, soll) -> bool:
    ist = None if ist == "" else ist
    soll = None if soll == "" else soll
    if isinstance(ist, bool) or isinstance(soll, bool):
        return ist == soll
    if isinstance(ist, (int, float)) and isinstance(soll, (int, float)):
        return abs(ist - soll) <= max(TOLERANZ, abs(soll) * 1e-9)
    return ist == soll


def vergleichen(wb, modell: Modell, js: dict, leer_kst=()) -> list:
    """Liste der Abweichungen (Bereich, Zeile, Excel, Browser)."""
    abw = []
    leer_kst = {tuple(x) for x in leer_kst}
    n = prognosejahre()
    kap = modell.kapazitaet

    def pruefe(name, zeile, excel, browser):
        if (zeile, name) in leer_kst and excel == 0 and browser == "":
            return
        if not gleich(excel, browser):
            abw.append((name, zeile, excel, browser))

    def je_zeile(name, werte_js):
        excel = spalte(wb, name)
        for z, w in werte_js:
            pruefe(name, z, excel[z], w)

    # Objekte
    objekte = js["objekte"]
    for f in OBJEKT_FELDER:
        je_zeile(f.name, [(j, o["felder"].get(f.key)) for j, o in enumerate(objekte) if o])
    je_zeile(STATUS_NAME, [(j, o["status"]) for j, o in enumerate(objekte) if o])
    je_zeile("obj_Annahmen", [(j, o["annahmen"]) for j, o in enumerate(objekte) if o])
    for s in OBJEKT_ANLAGEN_SPALTEN:
        je_zeile(s.name, [(j, o["anl"].get(s.key)) for j, o in enumerate(objekte) if o])
    # Anlagen
    anlagen = js["anlagen"]
    if anlagen:
        je_zeile("anl_AfA", [(j, a["afa"]) for j, a in enumerate(anlagen) if a])
        for s in ANLAGE_SPALTEN:
            je_zeile(s.name, [(j, a[s.key]) for j, a in enumerate(anlagen) if a])
        je_zeile(ANLAGE_STATUS_NAME, [(j, a["status"]) for j, a in enumerate(anlagen) if a])
        jahre = bereich(wb, ANLAGE_JAHRE_NAME)
        for j, a in enumerate(anlagen):
            if a:
                for i in range(n):
                    pruefe(ANLAGE_JAHRE_NAME, f"{j}/{i}", jahre[j][i], a["jahre"][i])
    # Prognose: je Objekt und Neuobjekt ein Block
    bloecke = [(j, s) for j, s in enumerate(js["prognose"]["bestand"]) if s]
    bloecke += [(kap.objekte + r, s) for r, s in enumerate(js["prognose"]["neu"]) if s]
    for sp in PROGNOSE_SPALTEN:
        excel = spalte(wb, sp.name)
        for block, s in bloecke:
            for i in range(n):
                pruefe(sp.name, block * n + i, excel[block * n + i], s[sp.key][i])
    # Verkäufe und Rücklagen (gleiche Zeilen)
    vk = js["verkaeufe"]
    for f in VERKAUF_FELDER:
        if f.key != "bezeichnung":
            je_zeile(f.name, [(j, v["felder"].get(f.key)) for j, v in enumerate(vk)])
    for s in VERKAUF_SPALTEN:
        je_zeile(s.name, [(j, v["berechnet"].get(s.key)) for j, v in enumerate(vk)])
    je_zeile(VERKAUF_STATUS_NAME, [(j, v["status"]) for j, v in enumerate(vk)])
    rl = js["ruecklagen"]
    for s in RUECKLAGE_SPALTEN:
        je_zeile(s.name, [(j, (r or {}).get(s.key)) for j, r in enumerate(rl)])
    for s in RUECKLAGE_JAHR_SPALTEN:
        je_zeile(s.name, [(i, w) for i, w in enumerate(js["spiegel"][s.key])])
    # Neuobjekte
    neu = js["neuobjekte"]
    for f in NEU_FELDER:
        je_zeile(f.name, [(r, x["felder"].get(f.key)) for r, x in enumerate(neu)
                          if x["felder"].get("neu_id")])
    quell_keys = {f"{feld}_{key}": (k, key) for k, (feld, _, _) in enumerate(QUELLEN)
                  for key in ("rl_geb", "rl_gub", "ue1", "ue2", "ue3", "erloes")}
    for s in NEU_SPALTEN:
        def wert_neu(x, key=s.key):
            b = x["berechnet"]
            if key in quell_keys:
                k, sub = quell_keys[key]
                return b["quellen"][k][sub]
            return b.get(key)
        je_zeile(s.name, [(r, wert_neu(x)) for r, x in enumerate(neu) if x["berechnet"]])
    je_zeile(NEU_STATUS_NAME, [(r, x["status"]) for r, x in enumerate(neu) if x["felder"].get("neu_id")])
    # Darlehen: oben Hauptkredite, darunter Zweitkredite
    for r, d in enumerate(js["darlehen"]):
        if not d:
            continue
        for teil, zeile in (("haupt", r), ("zweit", kap.neuobjekte + r)):
            kredit = d[teil]
            if not kredit:
                pruefe("dl_ID", zeile, spalte(wb, "dl_ID")[zeile], None)
                continue
            for s in DARLEHEN_SPALTEN:
                pruefe(s.name, zeile, spalte(wb, s.name)[zeile], kredit.get(s.key))
            for key, _, name in DARLEHEN_BLOECKE:
                excel = bereich(wb, name)[zeile]
                for i in range(n):
                    pruefe(name, f"{zeile}/{i}", excel[i], kredit[key][i])
    # Liquidität, Auswertung, Vergleich
    for sz in SZENARIEN:
        werte = js["szenarien"][sz.key]
        for s in liq_spalten(sz):
            je_zeile(s.name, list(enumerate(werte["liq"][s.key])))
        for s in aus_spalten(sz):
            je_zeile(s.name, list(enumerate(werte["aus"][s.key])))
        je_zeile(f"vg_{sz.key}", [(k, z[sz.key]) for k, z in enumerate(js["vergleich"])])
    assert len(js["vergleich"]) == len(VERGLEICH_KENNZAHLEN)
    for name, key in (("vg_DiffB", "diffB"), ("vg_DiffC", "diffC"), ("vg_DiffBaseline", "diffBaseline")):
        je_zeile(name, [(k, z[key]) for k, z in enumerate(js["vergleich"])])
    anzahl = spalte(wb, "pr_Anzahl")
    for k, p in enumerate(PRUEFUNGEN):
        if p.key not in OHNE_BROWSER:
            pruefe(f"pr_Anzahl:{p.key}", k, anzahl[k], js["pruefung"][p.key])
    return abw


# ---- Lauf ----

def faelle() -> list:
    """(Index, Name) der Fälle, die sich als Modell an das Rechenwerk geben lassen."""
    return [(i, f[0]) for i, f in enumerate(pruefen._faelle())
            if len(f) <= 3 and isinstance(f[1], (Modell, list, pruefen.Getippt))]


def modell_zu(i: int) -> tuple:
    """Modell des Falls wie in pruefen.lauf und die danach getippten Zellen. Getippt wird
    nur in Blättern, die das Rechenwerk nicht kennt (Kostenstellen, Ansichten) oder in
    Zellen, deren Wert der Abgleich aus der Mappe liest."""
    fall, objekte, *_ = pruefen._faelle()[i]
    zellen = None
    if isinstance(objekte, pruefen.Getippt):
        objekte, zellen = objekte.modell, objekte.zellen
    modell = objekte if isinstance(objekte, Modell) else Modell(objekte=objekte)
    modell = dataclasses.replace(modell)
    if not fall.startswith("Erhaltung"):
        modell.parameter = {**pruefen.OHNE_ALTERUNG, **modell.parameter}
    if fall not in pruefen.VOLLE_GROESSE:
        modell.kapazitaet = pruefen.TEST_KAPAZITAET
    return modell, zellen


def rechne_browser(eingaben: dict) -> dict:
    node = shutil.which("node")
    if not node:
        sys.exit("Node.js (node) nicht gefunden.")
    lauf = subprocess.run([node, str(CLI)], input=json.dumps(eingaben), capture_output=True,
                          text=True, check=True, timeout=120)
    return json.loads(lauf.stdout)


def lauf(i: int, ordner: Path, ausfuehrlich: bool) -> tuple:
    fall = pruefen._faelle()[i][0]
    modell, zellen = modell_zu(i)
    wb = pruefen.durchrechnen(modell, ordner / f"fall{i}", zellen, profil=ordner / "lo-profil")
    ein = eingabe(modell, wb)
    # zweiter Weg: Eingaben so, wie die Browser-Fassung sie aus der gespeicherten Mappe liest
    ein_mappe = eingabe_aus_mappe(wb.pfad)
    for j, name in ein["_leer_kst"]:   # wie oben: LibreOffice-Lesart leerer Bezüge
        key = next(f.key for f in OBJEKT_FELDER if f.name == name)
        ein_mappe["objekte"][j][key] = ""
        ein_mappe["objekte"][j].setdefault("herkunft", {})[key] = "annahme"
    js, js_mappe = rechne_browser([ein, ein_mappe])
    for ergebnis in (js, js_mappe):
        if "fehler" in ergebnis:
            return fall, [("Rechenwerk", 0, "", ergebnis["fehler"].splitlines()[0])], ergebnis["fehler"]
    abw = vergleichen(wb, modell, js, ein["_leer_kst"])
    abw += [(f"aus Mappe: {name}", zeile, excel, browser)
            for name, zeile, excel, browser in vergleichen(wb, modell, js_mappe, ein["_leer_kst"])]
    return fall, abw, ""


_ORDNER = None


def _start(tmp: str) -> None:
    global _ORDNER
    _ORDNER = Path(tempfile.mkdtemp(prefix=f"abgleich{os.getpid()}_", dir=tmp))
    pruefen._faelle()


def _im_arbeiter(i: int, ausfuehrlich: bool) -> tuple:
    return lauf(i, _ORDNER, ausfuehrlich)


def main() -> int:
    argumente = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    argumente.add_argument("filter", nargs="?", default="", help="Teil des Fallnamens")
    argumente.add_argument("-j", "--prozesse", type=int, default=os.cpu_count() or 1)
    argumente.add_argument("-v", "--ausfuehrlich", action="store_true",
                           help="alle Abweichungen statt der ersten zehn je Fall")
    args = argumente.parse_args()
    auswahl = [(i, name) for i, name in faelle() if args.filter in name]
    gesamt = 0
    abweichend = 0
    with tempfile.TemporaryDirectory() as tmp, concurrent.futures.ProcessPoolExecutor(
            max(1, min(args.prozesse, len(auswahl))), initializer=_start,
            initargs=(tmp,)) as pool:
        laeufe = {pool.submit(_im_arbeiter, i, args.ausfuehrlich): name for i, name in auswahl}
        for fertig, erledigt in enumerate(concurrent.futures.as_completed(laeufe), start=1):
            fall, abw, stapel = erledigt.result()
            gesamt += len(abw)
            abweichend += bool(abw)
            print(f"[{fertig}/{len(auswahl)}] {'ABWEICHUNG' if abw else 'gleich'}  {fall}"
                  + (f"  ({len(abw)})" if abw else ""), flush=True)
            for name, zeile, excel, browser in (abw if args.ausfuehrlich else abw[:10]):
                print(f"    {name}[{zeile}]  Excel {excel!r}  Browser {browser!r}")
            if stapel:
                print(stapel)
    print(f"\n{len(auswahl) - abweichend} von {len(auswahl)} Fällen gleich"
          + (f", {gesamt} Abweichung(en)" if gesamt else ""))
    return 1 if gesamt else 0


if __name__ == "__main__":
    sys.exit(main())
