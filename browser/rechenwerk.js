/*
 * Rechenwerk des Prognosemodells V+V im Browser.
 *
 * Eine reine Funktion: rechne(eingabe) -> ergebnis. Keine Abhängigkeiten, läuft im Browser
 * (window.Rechenwerk) und in Node (require), damit der Abgleich gegen die Excel-Mappe
 * (pruefung/abgleich.py) dieselbe Datei rechnet wie die HTML-Seite.
 *
 * Die Rechnung folgt Schritt für Schritt den Formeln aus prognosemodell/formeln.py. Die
 * Schlüssel der Ergebnisse sind dieselben wie die Spaltenschlüssel in modelle.py
 * (PROGNOSE_SPALTEN, VERKAUF_SPALTEN, ...), so lässt sich jede Zahl einem benannten Bereich
 * der Mappe zuordnen. Wer eine Formel in formeln.py ändert, ändert sie hier mit; der
 * Abgleich zeigt jede Abweichung.
 *
 * Eingabe (siehe browser/README.md):
 *   parameter: {Basisjahr: 2026, ...}  Namen ohne „par_“; berechnete Parameter fehlen
 *   objekte, verkaeufe, neuobjekte, anlagen: Listen von Feldern wie in modelle.py
 *     null/undefined = Feld leer gelassen (Annahme greift, wo es eine gibt), "" = leer
 *   afa_plan: {ObjektID: {Jahr: AfA}}, neukauf: [{kst, name, werte: {miete: {Jahr: Wert}}}]
 *   schnellcheck: true/false
 */
(function (wurzel, fabrik) {
  if (typeof module === "object" && module.exports) module.exports = fabrik();
  else wurzel.Rechenwerk = fabrik();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---- Konstanten (gleich modelle.py) ----
  const STATUS_OK = "OK";
  const STATUS_AUS = "ausgeschaltet";
  const STATUS_6B_UNZULAESSIG = "§ 6b unzulässig: Vorbesitzzeit zu kurz";
  const STATUS_ANNAHME_GELOESCHT = "Wert fehlt: Annahme gelöscht (Makro „Annahmen wiederherstellen“)";
  const ART_UNTERNEHMEN = "Unternehmen";
  const ART_GUB = "G+B", ART_GEBAEUDE = "Gebäude";
  const ARTEN_ABNUTZBAR = ["Gebäude", "BGA", "im Bau"];
  const GRUPPE_ABNUTZBAR = "abnutzbar";
  const ANL_NICHT_IM_MODELL = "nicht im Modell (Art)";
  const ANL_OHNE_OBJEKT = "ohne ObjektID";
  const ANL_OBJEKT_FEHLT = "ObjektID fehlt im Blatt Objekte";
  const STEUERWELT_GMBH = "GmbH", STEUERWELT_KG = "GmbH & Co. KG";
  const FIN_KREDIT = "Kredit";
  const TILGUNG_ANNUITAET = "Annuität", TILGUNG_LINEAR = "linear", TILGUNG_ENDFAELLIG = "endfällig";
  const UEBERNAHME_HAUPT = "auf Hauptkredit";
  const QUELLEN = ["quelle", "quelle2", "quelle3"];
  const VERGLEICH_JAHRE = 10;
  const SZENARIEN = ["A", "B", "C", "Baseline"];

  // Felder mit Annahme, die leer ein Fehler wären (Status „Annahme gelöscht“)
  const OBJEKT_ANNAHMEN = ["erhaltung", "verkehrswert", "vk_quote_gebaeude", "afa_satz", "kaufjahr",
    "ak_gebaeude", "ak_gub", "restbuchwert", "afa_jahr", "san_jahr", "san_betrag", "zinssatz", "rate"];
  // zählen als Annahme (blau), ohne die nur abgeleiteten (Steigerungen, AfA je Jahr)
  const OBJEKT_ANNAHMEN_ZAEHLEN = ["erhaltung", "verkehrswert", "vk_quote_gebaeude", "afa_satz",
    "kaufjahr", "ak_gebaeude", "ak_gub", "restbuchwert", "baujahr", "san_jahr", "san_betrag",
    "zinssatz", "rate"];
  const OBJEKT_KRITISCH = ["vk_quote_gebaeude", "afa_satz", "kaufjahr", "ak_gebaeude", "ak_gub",
    "restbuchwert"];
  const KST_VERKNUEPFT = ["miete", "erhaltung", "weitere_einnahmen", "weitere_ausgaben", "afa_bwa", "zinsen"];
  const NEU_PFLICHT = ["neu_id", "kaufjahr", "kaufpreis", "afa_satz"];

  // ---- Excel-Semantik ----
  const leer = (v) => v === null || v === undefined || v === "";
  const istZahl = (v) => typeof v === "number" && isFinite(v);
  /** N(): Zahl oder 0, wie Excel bei Text und leeren Zellen. */
  const N = (v) => (istZahl(v) ? v : v === true ? 1 : 0);
  /** Textvergleich wie Excel „=“: ohne Groß- und Kleinschreibung. */
  const gleich = (a, b) => String(leer(a) ? "" : a).toLowerCase() === String(leer(b) ? "" : b).toLowerCase();
  const text = (v) => (leer(v) ? "" : String(v));
  /** ROUND: kaufmännisch, halbe Werte von 0 weg. */
  function runden(x, stellen) {
    const f = Math.pow(10, stellen);
    return (Math.sign(x) * Math.round(Math.abs(x) * f * (1 + 1e-15))) / f;
  }
  /** ROUNDUP: von 0 weg auf stellen Stellen. */
  function aufrunden(x, stellen) {
    const f = Math.pow(10, stellen);
    return (Math.sign(x) * Math.ceil(Math.abs(x) * f - 1e-9)) / f;
  }
  /** TRIM: Leerzeichen außen weg, innen zu einem. */
  const trim = (s) => text(s).replace(/ +/g, " ").trim();
  const summe = (liste) => liste.reduce((a, b) => a + N(b), 0);
  const jahrAus = (datum) => {
    if (istZahl(datum)) return datum;     // schon ein Jahr
    const m = /^(\d{4})/.exec(text(datum));
    return m ? Number(m[1]) : null;
  };

  // ---- Parameter ----
  function parameterBerechnen(roh) {
    const p = Object.assign({}, roh);
    const def = (k, v) => { if (!istZahl(p[k]) && !(typeof p[k] === "string" && p[k] !== "")) p[k] = v; };
    def("Startjahr", N(p.Basisjahr) + 1);
    def("Endjahr", N(p.Basisjahr) + N(p.Prognosejahre));
    def("AnlStand", N(p.Basisjahr) - 1);
    def("KGGewSt", gleich(p.KGKuerzung, "ja") ? 0 : N(p.KGMesszahl) * N(p.KGHebesatz));
    def("SteuersatzEff", gleich(p.Steuerwelt, STEUERWELT_KG)
      ? p.KGGewSt + (N(p.KGESt) - Math.min(N(p.KGAnrechnung) * N(p.KGMesszahl), p.KGGewSt, N(p.KGESt))) * (1 + N(p.KGSoli))
      : N(p.Steuersatz));
    def("VerlustSatz", gleich(p.Steuerwelt, STEUERWELT_KG) && gleich(p.KGVerlustausgleich, "ja")
      ? N(p.KGESt) * (1 + N(p.KGSoli)) : 0);
    return p;
  }

  // ---- Anlagen (Blatt Anlagen) ----
  function anlagenRechnen(eingabe, P, objektIds, jahre) {
    const liste = (eingabe.anlagen || []).map((a) => Object.assign({}, a));
    const nummern = liste.map((a) => a.nr);
    return liste.map((a) => {
      if (leer(a.nr)) return null;
      const methode = text(a.methode);
      const degressiv = gleich(methode, "degressiv"), keine = gleich(methode, "keine");
      const bw = N(a.bw_stand), satz = N(a.satz);
      const z = { felder: a, herkunft: {} };
      if (a.afa === null || a.afa === undefined) {
        z.afa = degressiv ? runden(bw * satz, 2) : keine ? 0 : aufrunden(runden(N(a.ahk) * satz, 2), 0);
        z.herkunft.afa = "annahme";
      } else z.afa = a.afa;
      const afa = N(z.afa);
      const buchwert = (jahr) => {
        const n = Math.max(jahr - P.AnlStand, 0);
        return degressiv ? bw * Math.pow(1 - satz, n) : keine ? bw : Math.max(bw - afa * n, 0);
      };
      const afaJahr = (jahr) => {
        if (jahr <= P.AnlStand) return 0;
        const n = jahr - 1 - P.AnlStand;
        return degressiv ? bw * Math.pow(1 - satz, n) * satz : keine ? 0 : Math.min(afa, Math.max(bw - afa * n, 0));
      };
      const art = text(a.art);
      z.gruppe = gleich(art, ART_GUB) ? ART_GUB : ARTEN_ABNUTZBAR.some((x) => gleich(art, x)) ? GRUPPE_ABNUTZBAR : "";
      const jahrZugang = jahrAus(a.datum);
      z.zugang = (gleich(art, ART_GUB) || gleich(art, ART_GEBAEUDE)) && N(jahrZugang) > 0 ? jahrZugang : "";
      z.bw_basis = buchwert(P.Basisjahr);
      z.afa_basis = P.AnlStand === P.Basisjahr ? N(a.afa_stand) : afaJahr(P.Basisjahr);
      const pflicht = ["nr", "art", "bw_stand", "methode"].filter((k) => !leer(a[k])).length;
      z.status = pflicht < 4 ? "Pflichtfeld fehlt"
        : nummern.filter((n) => gleich(n, a.nr)).length > 1 ? "Inventar-Nr. doppelt"
        : z.gruppe === "" ? ANL_NICHT_IM_MODELL
        : leer(a.objekt_id) ? ANL_OHNE_OBJEKT
        : !objektIds.some((id) => gleich(id, a.objekt_id)) ? ANL_OBJEKT_FEHLT : STATUS_OK;
      z.jahre = jahre.map(afaJahr);
      return z;
    });
  }

  /** Anlagen eines Objekts mit Status OK, optional gefiltert. */
  function anlagenVon(anlagen, id, filter) {
    return anlagen.filter((a) => a && a.status === STATUS_OK && gleich(a.felder.objekt_id, id)
      && (!filter || filter(a)));
  }

  // ---- Objekte (Blatt Objekte mit Auffülllogik) ----
  function objekteRechnen(eingabe, P, anlagen) {
    const ids = (eingabe.objekte || []).map((o) => o.objekt_id);
    return (eingabe.objekte || []).map((roh) => {
      const o = { felder: {}, herkunft: {}, anl: {} };
      const f = o.felder;
      if (leer(roh.objekt_id)) return null;
      const id = roh.objekt_id;
      // Hilfsspalten aus dem Blatt Anlagen
      const abn = anlagenVon(anlagen, id, (a) => a.gruppe === GRUPPE_ABNUTZBAR);
      const gebaeude = anlagenVon(anlagen, id, (a) => gleich(a.felder.art, ART_GEBAEUDE) && N(a.zugang) > 0);
      o.anl = {
        anl_abn: abn.length,
        anl_ak: abn.filter((a) => N(a.felder.ahk) > 0).length,
        anl_gub: anlagenVon(anlagen, id, (a) => a.gruppe === ART_GUB).length,
        anl_kauf: anlagenVon(anlagen, id, (a) => N(a.zugang) > 0).length,
        anl_bau: gebaeude.length ? Math.min(...gebaeude.map((a) => a.zugang)) : "",
      };
      o.anl.anl_afa = o.anl.anl_abn === 0 ? "" : summe(abn.map((a) => a.afa_basis));
      const deckt = {
        ak_gebaeude: o.anl.anl_abn > 0 && o.anl.anl_ak === o.anl.anl_abn,
        afa_satz: o.anl.anl_abn > 0 && o.anl.anl_ak === o.anl.anl_abn,
        ak_gub: o.anl.anl_gub > 0,
        kaufjahr: o.anl.anl_kauf > 0,
        baujahr: N(o.anl.anl_bau) > 0,
        restbuchwert: o.anl.anl_abn > 0,
        afa_jahr: o.anl.anl_abn > 0,
      };
      const ahk = summe(abn.map((a) => a.felder.ahk));
      const ausAnlagen = {
        ak_gebaeude: () => ahk,
        afa_satz: () => (ahk > 0 ? summe(abn.map((a) => a.afa)) / ahk : null),
        ak_gub: () => summe(anlagenVon(anlagen, id, (a) => a.gruppe === ART_GUB).map((a) => a.bw_basis)),
        kaufjahr: () => Math.min(...anlagenVon(anlagen, id, (a) => N(a.zugang) > 0).map((a) => a.zugang)),
        baujahr: () => o.anl.anl_bau,
        restbuchwert: () => summe(abn.map((a) => a.bw_basis)),
        afa_jahr: () => summe(abn.map((a) => a.jahre[0])),
      };
      const n = (k) => N(f[k]);
      const annahme = {
        erhaltung: () => n("miete") * P.AnnErhQuote,
        mietsteig: () => P.Mietsteig,
        erhaltsteig: () => P.Erhaltsteig,
        verkehrswert: () => (gleich(f.art, ART_UNTERNEHMEN) ? 0 : (n("miete") + n("weitere_einnahmen")) * P.AnnVervielfaeltiger),
        vk_quote_gebaeude: () => P.AnnGebaeudeanteil,
        afa_satz: () => P.AnnAfASatz,
        kaufjahr: () => P.Basisjahr - P.AnnHaltedauer,
        ak_gebaeude: () => (n("afa_bwa") > 0 && n("afa_satz") > 0 ? n("afa_bwa") / n("afa_satz")
          : n("verkehrswert") * n("vk_quote_gebaeude") / Math.pow(1 + P.Wertsteig, P.Basisjahr - n("kaufjahr"))),
        ak_gub: () => (n("vk_quote_gebaeude") > 0 ? n("ak_gebaeude") * (1 - n("vk_quote_gebaeude")) / n("vk_quote_gebaeude") : 0),
        restbuchwert: () => (!leer(f.afa_bwa) && n("afa_bwa") === 0 ? 0
          : Math.max(n("ak_gebaeude") * (1 - (P.Basisjahr - n("kaufjahr") + 1) * n("afa_satz")), 0)),
        afa_jahr: () => (!leer(f.afa_bwa) ? n("afa_bwa") : n("ak_gebaeude") * n("afa_satz")),
        baujahr: () => "",
        san_jahr: () => {
          if (P.SanQuote === 0 || n("baujahr") === 0) return 0;
          const j = Math.max(P.Startjahr + P.SanVorlauf, n("baujahr") + P.SanAlter);
          return j > P.Endjahr ? 0 : j;
        },
        san_betrag: () => (n("san_jahr") === 0 ? 0 : n("verkehrswert") * n("vk_quote_gebaeude") * P.SanQuote),
        zinssatz: () => (n("restschuld") > 0 ? n("zinsen") / n("restschuld") : 0),
        rate: () => (n("restschuld") > 0 ? n("restschuld") * (n("zinssatz") + P.AnnDarlTilgung) : 0),
      };
      const reihenfolge = ["objekt_id", "name", "art", "miete", "erhaltung", "weitere_einnahmen",
        "weitere_ausgaben", "afa_bwa", "zinsen", "mietsteig", "erhaltsteig", "miete_fix_ab",
        "verkehrswert", "vk_quote_gebaeude", "afa_satz", "kaufjahr", "ak_gebaeude", "ak_gub",
        "restbuchwert", "afa_jahr", "baujahr", "san_jahr", "san_betrag", "restschuld", "zinssatz", "rate"];
      for (const k of reihenfolge) {
        const v = roh[k];
        if (v !== null && v !== undefined) {
          f[k] = v;
          o.herkunft[k] = (roh.herkunft && roh.herkunft[k]) || (v === "" ? "leer" : "eingabe");
        } else if (annahme[k]) {
          if (deckt[k]) {
            const w = ausAnlagen[k]();
            f[k] = w === null ? annahme[k]() : w;
            o.herkunft[k] = "anlagen";
          } else {
            f[k] = annahme[k]();
            o.herkunft[k] = "annahme";
          }
        } else {
          f[k] = "";
          o.herkunft[k] = "leer";
        }
      }
      // Status
      const pflicht = !leer(f.objekt_id) && !leer(f.miete);
      const geloescht = OBJEKT_ANNAHMEN.some((k) => leer(f[k]));
      o.status = !pflicht ? "Pflichtfeld fehlt"
        : geloescht ? STATUS_ANNAHME_GELOESCHT
        : ids.filter((x) => gleich(x, id)).length > 1 ? "ObjektID doppelt"
        : N(f.kaufjahr) > P.Basisjahr ? "Kaufjahr nach Basisjahr"
        : N(f.restbuchwert) > N(f.ak_gebaeude) ? "Restbuchwert > AK Gebäude" : STATUS_OK;
      // Anzahl Annahmen (blau) und kritische Annahmen; eingelesene Werte zählen nicht
      const istAnnahme = (k) => {
        if (o.herkunft[k] !== "annahme") return false;
        if (k === "zinssatz" || k === "rate") return n("restschuld") > 0;
        if (k === "baujahr") return !leer(f.baujahr);
        return true;
      };
      o.annahmen = OBJEKT_ANNAHMEN_ZAEHLEN.filter(istAnnahme).length;
      o.kritischOhneVerkauf = OBJEKT_KRITISCH.filter(istAnnahme).length;
      o.anl.anl_bau_offen = o.herkunft.baujahr === "anlagen" && N(o.anl.anl_bau) > 0 ? 1 : 0;
      o.anl.anl_diff = o.anl.anl_afa === "" || leer(f.afa_bwa) ? "" : o.anl.anl_afa - n("afa_bwa");
      return o;
    });
  }

  // ---- Verkäufe, Teil 1: Umfang, Annahmen (vor der Prognose) ----
  function umfang(id, weitere) {
    if (leer(id)) return "";
    let liste = trim(weitere).split(",").join(";");
    for (const alt of ["; ", " ;", ";;"]) liste = liste.split(alt).join(";");
    return ";" + id + ";" + (trim(weitere) === "" ? "" : liste + ";");
  }
  const imPaket = (umf, id) => !leer(id) && text(umf).toLowerCase().includes((";" + id + ";").toLowerCase());

  function verkaeufeVorbereiten(eingabe, P, objekte) {
    const objNach = (id) => objekte.find((o) => o && gleich(o.felder.objekt_id, id));
    return (eingabe.verkaeufe || []).map((roh) => {
      const v = { felder: Object.assign({}, roh), herkunft: {} };
      const f = v.felder;
      for (const k of Object.keys(f)) if (f[k] === null || f[k] === undefined) f[k] = "";
      v.umfang = umfang(f.objekt_id, f.weitere_ids);
      if (eingabe.schnellcheck) {
        for (const k of ["nutzung_6b", "reinvest"]) {
          if (roh[k] === null || roh[k] === undefined) {
            f[k] = leer(f.objekt_id) ? "" : "ja";
            v.herkunft[k] = "annahme";
          }
        }
      }
      // Bezeichnung: Objektname (das Kostenstellenblatt gibt es im Browser nicht)
      const o = objNach(f.objekt_id);
      f.bezeichnung = o ? text(o.felder.name) : "";
      if (roh.preis === null || roh.preis === undefined) {
        v.herkunft.preis = "annahme";
        if (leer(f.objekt_id) || leer(f.jahr)) f.preis = "";
        else {
          const imV = objekte.filter((ob) => ob && imPaket(v.umfang, ob.felder.objekt_id));
          const vw = trim(f.weitere_ids) === ""
            ? (o ? N(o.felder.verkehrswert) : null)
            : summe(imV.map((ob) => N(ob.felder.verkehrswert)));
          f.preis = vw === null ? "" : vw * Math.pow(1 + P.Wertsteig, N(f.jahr) - P.Basisjahr);
        }
      }
      v.aktiv = !gleich(f.aktiv, "nein");
      return v;
    });
  }

  /** Verkaufsjahr eines Objekts (aktive Verkäufe mit Jahr, auch als weitere Kostenstelle);
   * 9999 ohne Verkauf. Dazu die Zeile des Verkaufs (1 = erste), 0 ohne. */
  function verkaufsjahr(verkaeufe, id) {
    const treffer = verkaeufe.map((v, i) => ({ v, i })).filter(({ v }) =>
      v.aktiv && imPaket(v.umfang, id) && !leer(v.felder.jahr));
    if (!treffer.length) return { jahr: 9999, nr: 0 };
    const jahr = Math.min(...treffer.map(({ v }) => N(v.felder.jahr)));
    const nr = Math.max(0, ...treffer.filter(({ v }) => N(v.felder.jahr) === jahr).map(({ i }) => i + 1));
    return { jahr, nr };
  }

  // ---- Prognose der Bestandsobjekte ----
  function prognoseBestand(o, P, jahre, anlagen, verkaeufe, afaPlan) {
    const f = o.felder;
    const n = (k) => N(f[k]);
    const { jahr: vj, nr: vkNr } = verkaufsjahr(verkaeufe, f.objekt_id);
    const mietsatz = leer(f.mietsteig) ? P.Mietsteig : n("mietsteig");
    const erhsatz = leer(f.erhaltsteig) ? P.Erhaltsteig : n("erhaltsteig");
    const abn = anlagenVon(anlagen, f.objekt_id, (a) => a.gruppe === GRUPPE_ABNUTZBAR);
    const plan = (afaPlan || {})[f.objekt_id] || {};
    const fix = n("miete_fix_ab");
    const ueber = (jahr) => (n("baujahr") === 0 ? 0 : Math.max(jahr - n("baujahr") - P.ErhAlterungAb, 0));
    const rs0 = n("restschuld"), zsatz = n("zinssatz"), rate = n("rate"), mit = rs0 > 0;
    const s = {};
    const neu = (k) => (s[k] = new Array(jahre.length));
    ["id", "jahr", "aktiv", "miete", "einnahmen", "erhaltung", "ausgaben", "afa", "afa_qng", "buchwert",
      "ergebnis", "verkehrswert", "bestand", "neu", "buchwert_gub", "buchwert_halten", "mit_quelle",
      "afa_ohne6b", "afa_qng_ohne6b", "buchwert_ohne6b", "buchwert_gub_ohne6b", "erhaltung_halten",
      "miete_halten", "einnahmen_halten", "quelle", "afa_halten", "kredit_zins", "tilgung", "abloesung",
      "restschuld", "zins_halten", "tilgung_halten", "restschuld_halten", "ergebnis_halten", "wirkung",
      "verkauf_nr", "unternehmen"].forEach(neu);
    jahre.forEach((t, i) => {
      const dt = t - P.Basisjahr;
      s.id[i] = f.objekt_id;
      s.jahr[i] = t;
      const aktiv = t <= vj ? 1 : 0;
      s.aktiv[i] = aktiv;
      const bis = Math.max((fix > 0 ? Math.min(t, fix - 1) : t) - P.Basisjahr, 0);
      s.miete_halten[i] = n("miete") * Math.pow(1 + mietsatz, bis);
      s.einnahmen_halten[i] = n("weitere_einnahmen") * Math.pow(1 + mietsatz, dt);
      const alterung = Math.pow(1 + P.ErhAlterung, ueber(t) - ueber(P.Basisjahr));
      const gross = t === n("san_jahr") ? n("san_betrag") : 0;
      s.erhaltung_halten[i] = (n("erhaltung") * alterung + gross) * Math.pow(1 + erhsatz, dt);
      s.miete[i] = s.miete_halten[i] * aktiv;
      s.einnahmen[i] = s.einnahmen_halten[i] * aktiv;
      s.erhaltung[i] = s.erhaltung_halten[i] * aktiv;
      const ausgabenHalten = n("weitere_ausgaben") * Math.pow(1 + P.Kostensteig, dt);
      s.ausgaben[i] = ausgabenHalten * aktiv;
      const bwHaltenVor = i === 0 ? n("restbuchwert") : s.buchwert_halten[i - 1];
      const ausPlan = plan[t] !== undefined ? plan[t] : plan[String(t)];
      const fortschreibung = o.anl.anl_abn > 0 ? summe(abn.map((a) => a.jahre[i])) : n("afa_jahr");
      s.afa_halten[i] = Math.min(istZahl(ausPlan) ? ausPlan : fortschreibung, bwHaltenVor);
      s.buchwert_halten[i] = Math.max(bwHaltenVor - s.afa_halten[i], 0);
      s.afa[i] = s.afa_halten[i] * aktiv;
      const bwVor = i === 0 ? n("restbuchwert") : s.buchwert[i - 1];
      s.buchwert[i] = Math.max(bwVor - s.afa[i], 0);
      s.ergebnis[i] = (s.miete[i] + s.einnahmen[i] - s.erhaltung[i] - s.ausgaben[i] - s.afa[i]) * aktiv;
      s.verkehrswert[i] = n("verkehrswert") * Math.pow(1 + P.Wertsteig, dt);
      s.bestand[i] = t < vj ? 1 : 0;
      s.neu[i] = 0;
      s.buchwert_gub[i] = n("ak_gub");
      s.mit_quelle[i] = 0;
      s.afa_ohne6b[i] = s.afa[i];
      s.afa_qng[i] = 0;
      s.afa_qng_ohne6b[i] = 0;
      s.buchwert_ohne6b[i] = s.buchwert[i];
      s.buchwert_gub_ohne6b[i] = s.buchwert_gub[i];
      s.quelle[i] = "";
      // Darlehen im Plan (Ablösung im Verkaufsjahr) und bei Halten
      const grob = n("zinsen") * Math.pow(1 + P.ZinsVeraenderung, dt);
      const vor = i === 0 ? rs0 : s.restschuld[i - 1];
      s.kredit_zins[i] = mit ? vor * zsatz : grob * aktiv;
      const regulaer = Math.min(Math.max(rate - s.kredit_zins[i], 0), vor);
      s.tilgung[i] = mit ? (s.bestand[i] === 0 ? vor : regulaer) : 0;
      s.restschuld[i] = mit ? vor - s.tilgung[i] : 0;
      s.abloesung[i] = mit ? s.tilgung[i] - regulaer : 0;
      const vorH = i === 0 ? rs0 : s.restschuld_halten[i - 1];
      s.zins_halten[i] = mit ? vorH * zsatz : grob;
      s.tilgung_halten[i] = mit ? Math.min(Math.max(rate - s.zins_halten[i], 0), vorH) : 0;
      s.restschuld_halten[i] = mit ? vorH - s.tilgung_halten[i] : 0;
      s.ergebnis_halten[i] = s.miete_halten[i] + s.einnahmen_halten[i] - s.erhaltung_halten[i]
        - ausgabenHalten - s.afa_halten[i];
      s.wirkung[i] = s.ergebnis[i] - s.ergebnis_halten[i] - s.kredit_zins[i] + s.zins_halten[i];
      s.verkauf_nr[i] = vj === 9999 ? 0 : vkNr;
      s.unternehmen[i] = gleich(f.art, ART_UNTERNEHMEN) ? 1 : 0;
    });
    return s;
  }

  // ---- Verkäufe, Teil 2: Buchwert, Aufteilung, Gewinn ----
  function verkaeufeRechnen(verkaeufe, objekte, bestand, P, jahre) {
    const prgWert = (k, id, jahr) => {
      let sum = 0;
      objekte.forEach((o, j) => {
        if (!o || !gleich(o.felder.objekt_id, id)) return;
        const i = jahre.indexOf(jahr);
        if (i >= 0) sum += N(bestand[j][k][i]);
      });
      return sum;
    };
    const durchschnitt = (k, ids, jahr) => {
      if (leer(jahr)) return "";
      const bisJ = Math.min(N(jahr) + VERGLEICH_JAHRE, P.Endjahr);
      if (bisJ <= N(jahr)) return "";
      let sum = 0;
      for (let t = N(jahr) + 1; t <= bisJ; t++) for (const id of ids) sum += prgWert(k, id, t);
      return sum / (bisJ - N(jahr));
    };
    verkaeufe.forEach((v, zeile) => {
      const f = v.felder;
      const o = objekte.find((ob) => ob && gleich(ob.felder.objekt_id, f.objekt_id));
      v.berechnet = {};
      const b = v.berechnet;
      b.umfang = v.umfang;
      if (leer(f.objekt_id) || !o) {
        v.status = leer(f.objekt_id) ? "" : null;   // Status unten
      }
      const imV = objekte.filter((ob) => ob && imPaket(v.umfang, ob.felder.objekt_id));
      const paket = trim(f.weitere_ids) !== "";
      const jahr = f.jahr;
      if (o) {
        const kaufjahr = paket ? Math.max(0, ...imV.map((ob) => N(ob.felder.kaufjahr))) : N(o.felder.kaufjahr);
        b.vorbesitz = N(jahr) - kaufjahr;
        b.buchwert_geb = summe((paket ? imV : [o]).map((ob) => prgWert("buchwert", ob.felder.objekt_id, N(jahr))));
        b.ak_gub = paket ? summe(imV.map((ob) => ob.felder.ak_gub)) : o.felder.ak_gub;
        b.nettoerloes = N(f.preis) - N(f.kosten);
        let quoteGeb;
        if (!paket) quoteGeb = o.felder.vk_quote_gebaeude;
        else if (imV.some((ob) => leer(ob.felder.vk_quote_gebaeude))) quoteGeb = "";
        else {
          const vw = summe(imV.map((ob) => ob.felder.verkehrswert));
          quoteGeb = vw === 0 ? "" : summe(imV.map((ob) => N(ob.felder.verkehrswert) * N(ob.felder.vk_quote_gebaeude))) / vw;
        }
        b.quote_gub = !leer(f.gub_betrag) ? (N(f.preis) > 0 ? N(f.gub_betrag) / f.preis : "")
          : !leer(f.anteil_gub) ? f.anteil_gub
          : leer(quoteGeb) ? "" : 1 - quoteGeb;
        const q = b.quote_gub;
        const mitQ = (fn) => (leer(q) ? "" : fn());
        b.erloes_gub = mitQ(() => b.nettoerloes * q);
        b.erloes_geb = mitQ(() => b.nettoerloes - b.erloes_gub);
        b.gewinn_geb = mitQ(() => b.erloes_geb - b.buchwert_geb);
        b.gewinn_gub = mitQ(() => b.erloes_gub - N(b.ak_gub));
        b.gewinn = mitQ(() => b.gewinn_geb + b.gewinn_gub);
        const ids = (paket ? imV : [o]).map((ob) => ob.felder.objekt_id);
        b.miete = leer(jahr) || N(jahr) < P.Startjahr || N(jahr) > P.Endjahr ? ""
          : summe(ids.map((id) => prgWert("miete", id, N(jahr))));
        b.rendite = leer(b.miete) || N(f.preis) <= 0 ? "" : b.miete / f.preis;
        b.faktor = N(b.miete) <= 0 || N(f.preis) <= 0 ? "" : f.preis / b.miete;
        b.miete_oe = durchschnitt("miete_halten", ids, jahr);
        b.erh_oe = durchschnitt("erhaltung_halten", ids, jahr);
        b.erh_quote = leer(b.erh_oe) || N(b.miete_oe) <= 0 ? "" : b.erh_oe / b.miete_oe;
      }
      // Status
      if (leer(f.objekt_id)) { v.status = ""; return; }
      const eintraege = (v.umfang.match(/;/g) || []).length - 1;
      const verkauftAnzahl = (id) => verkaeufe.filter((w) => w.aktiv && imPaket(w.umfang, id)).length;
      v.status = !v.aktiv ? STATUS_AUS
        : !o ? "ObjektID unbekannt"
        : eintraege > imV.length ? "Kostenstelle unbekannt: weitere Kostenstellen prüfen"
        : imV.some((ob) => verkauftAnzahl(ob.felder.objekt_id) > 1) ? "Objekt mehrfach verkauft"
        : leer(jahr) ? "Verkaufsjahr fehlt"
        : N(jahr) < P.Startjahr || N(jahr) > P.Endjahr ? "Verkaufsjahr außerhalb Raster"
        : leer(f.preis) ? "Verkaufspreis fehlt"
        : leer(b.quote_gub) ? "Aufteilung fehlt: Anteil G+B oder Verkehrswertanteil"
        : b.quote_gub > 1 ? "Preis G+B über Verkaufspreis"
        : gleich(f.nutzung_6b, "ja") && b.vorbesitz < P["6bVorbesitz"] ? STATUS_6B_UNZULAESSIG
        : STATUS_OK;
      v.gueltig = v.status === STATUS_OK || v.status === STATUS_6B_UNZULAESSIG;
      v.zeile = zeile;
    });
  }

  // ---- Rücklagen je Verkauf (ohne Übertragung) ----
  function ruecklagenRechnen(verkaeufe, P) {
    return verkaeufe.map((v) => {
      const f = v.felder;
      if (!(v.status === STATUS_OK && gleich(f.nutzung_6b, "ja"))) return null;
      const r = { id: "RL-" + f.objekt_id, objekt_id: f.objekt_id, jahr: f.jahr };
      r.geb = Math.max(N(v.berechnet.gewinn_geb), 0);
      r.gub = Math.max(N(v.berechnet.gewinn_gub), 0);
      r.betrag = r.geb + r.gub;
      r.fristjahr = N(f.jahr) + (gleich(f.neubau_6b, "ja") ? P["6bFristNeubau"] : P["6bFrist"]);
      return r;
    });
  }

  // ---- Neuobjekte ----
  function neuobjekteEingaben(eingabe, P, verkaeufe) {
    const erfasst = eingabe.neuobjekte || [];
    const zeilen = Math.max(erfasst.length, verkaeufe.length);
    const liste = [];
    for (let r = 0; r < zeilen; r++) {
      const roh = erfasst[r];
      const f = {}, herkunft = {};
      if (roh) {
        for (const [k, v] of Object.entries(roh)) f[k] = v === null || v === undefined ? "" : v;
      } else {
        // Reinvestition aus der Verkaufszeile gleicher Nummer
        const v = verkaeufe[r];
        const an = v && gleich(v.felder.reinvest, "ja") && v.gueltig
          && N(v.felder.jahr) + P.AnnReinvestJahre <= P.Endjahr;
        if (an) {
          Object.assign(f, {
            neu_id: "NEU-" + v.felder.objekt_id, name: "Reinvestition aus " + v.felder.objekt_id,
            kaufjahr: N(v.felder.jahr) + P.AnnReinvestJahre,
            kaufpreis: N(v.berechnet.nettoerloes) * P.AnnReinvestQuote / (1 + P.AnnNeuNebenkosten),
            quelle: gleich(v.felder.nutzung_6b, "ja") ? "RL-" + v.felder.objekt_id : "",
            erloes_fin: "ja",
          });
          for (const k of ["neu_id", "name", "kaufjahr", "kaufpreis", "quelle", "erloes_fin"]) herkunft[k] = "annahme";
        } else f.neu_id = "";
      }
      // Annahmen des Parameterblatts, nur in Zeilen ohne erfasstes Neuobjekt
      if (!roh && !leer(f.neu_id)) {
        const ann = { anteil_gub: P.AnnNeuAnteilGuB, afa_satz: P.AnnNeuAfASatz,
          afa_methode: P.AnnNeuAfAMethode, mietrendite: P.AnnNeuMietrendite,
          erhaltungsquote: P.AnnNeuErhQuote };
        for (const [k, w] of Object.entries(ann)) { f[k] = w; herkunft[k] = "annahme"; }
        f.nebenkosten = N(f.kaufpreis) * P.AnnNeuNebenkosten;
        herkunft.nebenkosten = "annahme";
      }
      for (const k of ["aktiv", "name", "kaufjahr", "kaufpreis", "anteil_gub", "nebenkosten", "afa_satz",
        "afa_methode", "qng", "qng_grenze", "mietrendite", "erhaltungsquote", "flaeche", "miete_qm",
        "mietsteig", "miet_warte", "miet_bis", "quelle", "quelle2", "quelle3", "kst", "erloes_fin",
        "fin_art", "kredit_betrag", "kredit_quote", "kredit_zins", "tilgungsart", "tilgung", "laufzeit",
        "tilgungsfrei", "kredit2_anteil", "kredit2_betrag", "kredit2_zins", "kredit2_tilgungsart",
        "kredit2_tilgung", "kredit2_laufzeit", "kredit2_tilgungsfrei", "kredit2_uebernahme"]) {
        if (f[k] === undefined) f[k] = "";
      }
      // AK G+B und AK Gebäude: vorbelegt aus dem Anteil G+B, eingetippte Beträge gehen vor
      const ak = N(f.kaufpreis) + N(f.nebenkosten);
      const rohGub = roh ? roh.ak_gub_eingabe : null, rohGeb = roh ? roh.ak_geb_eingabe : null;
      if (rohGub === null || rohGub === undefined) {
        f.ak_gub_eingabe = leer(f.neu_id) || leer(f.anteil_gub) ? "" : ak * f.anteil_gub;
        herkunft.ak_gub_eingabe = "abgeleitet";
      } else f.ak_gub_eingabe = rohGub;
      if (rohGeb === null || rohGeb === undefined) {
        f.ak_geb_eingabe = leer(f.neu_id) || leer(f.ak_gub_eingabe) ? "" : ak - f.ak_gub_eingabe;
        herkunft.ak_geb_eingabe = "abgeleitet";
      } else f.ak_geb_eingabe = rohGeb;
      liste.push({ felder: f, herkunft, erfasst: !!roh });
    }
    return liste;
  }

  function neuobjekteRechnen(neu, ruecklagen, verkaeufe, objekte, P, neukaufFort) {
    const rlIndex = (q) => ruecklagen.findIndex((r) => r && gleich(r.id, q));
    const rl = (q) => ruecklagen[rlIndex(q)];
    const VORJAHR = gleich(P["6bVorjahr"], "ja") ? 1 : 0;
    const neuIds = neu.map((x) => x.felder.neu_id);
    const objIds = objekte.filter(Boolean).map((o) => o.felder.objekt_id);
    neu.forEach((x, r) => {
      const f = x.felder;
      if (leer(f.neu_id)) { x.berechnet = null; x.status = ""; return; }
      const b = (x.berechnet = {});
      const kj = f.kaufjahr;
      const ak = N(f.kaufpreis) + N(f.nebenkosten);
      const pflicht = NEU_PFLICHT.every((k) => !leer(f[k]));
      const aufteilung = !leer(f.anteil_gub) || !leer(f.ak_gub_eingabe);
      b.gueltig = !gleich(f.aktiv, "nein") && pflicht && aufteilung
        && neuIds.filter((id) => gleich(id, f.neu_id)).length === 1
        && !objIds.some((id) => gleich(id, f.neu_id))
        && N(kj) >= P.Startjahr && N(kj) <= P.Endjahr ? 1 : 0;
      const gueltig = b.gueltig === 1;
      b.ak_gub_neu = leer(f.ak_gub_eingabe) ? ak * N(f.anteil_gub) : f.ak_gub_eingabe;
      b.ak_geb_neu = leer(f.ak_geb_eingabe) ? ak - b.ak_gub_neu : f.ak_geb_eingabe;
      b.anteil_gub_verw = ak > 0 ? b.ak_gub_neu / ak : "";
      b.ak_abweichung = b.ak_gub_neu + b.ak_geb_neu - ak;
      const akGeb = b.ak_geb_neu, akGub = b.ak_gub_neu;
      const kreditFest = !gleich(f.fin_art, FIN_KREDIT) ? 0
        : !leer(f.kredit_betrag) ? N(f.kredit_betrag) : ak * N(f.kredit_quote);
      const fl = N(f.flaeche);
      const ueberObergrenze = fl > 0 ? akGeb / fl > P.QNGObergrenzeQm : false;
      const qngBasis = (basis) => (gleich(f.qng, "ja") && gueltig
        ? (ueberObergrenze ? 0 : Math.max(fl > 0 ? Math.min(basis, P.QNGHoechstQm * fl) : basis, 0)) : 0);
      b.qng_basis_ohne6b = qngBasis(akGeb);
      const qngKaufjahr = P.QNGJahre > 0 ? N(b.qng_basis_ohne6b) * P.QNGSatz : 0;
      const qngGrenze = gleich(f.qng_grenze, "ja") && gleich(f.qng, "ja") && !ueberObergrenze
        ? Math.max(akGeb - (fl > 0 ? P.QNGHoechstQm * fl : akGeb), 0) : akGeb;
      // je Quelle
      const qs = QUELLEN.map((feld, k) => {
        const q = f[feld];
        const doppelt = QUELLEN.slice(0, k).some((g) => gleich(f[g], q));
        const bekannt = gueltig && !leer(q) && rlIndex(q) >= 0 && !doppelt;
        const r0 = bekannt ? rl(q) : null;
        return {
          feld, q, bekannt,
          moeglich: bekannt && N(kj) >= N(r0.jahr) - VORJAHR && N(kj) <= N(r0.fristjahr),
          vorjahr: bekannt && N(r0.jahr) === N(kj) + 1,
          erloesMoeglich: bekannt && N(kj) >= N(r0.jahr),
          aus: rlIndex(q) < 0 && gleich(text(q).slice(0, 3), "RL-")
            && verkaeufe.some((v) => gleich(v.felder.objekt_id, text(q).slice(3)) && gleich(v.felder.aktiv, "nein")),
        };
      });
      // Was die Zeilen darüber aus der Rücklage q genommen haben (über alle Quellspalten)
      const verbraucht = (key, q) => {
        let sum = 0;
        for (let j = 0; j < r; j++) {
          const y = neu[j];
          if (!y.berechnet) continue;
          QUELLEN.forEach((g, k) => { if (gleich(y.felder[g], q)) sum += N(y.berechnet.quellen[k][key]); });
        }
        return sum;
      };
      const gebGrenze = (k) => Math.min(akGeb - (qs[k].vorjahr ? qngKaufjahr : 0), qngGrenze);
      const quellen = qs.map(() => ({}));
      qs.forEach((s, k) => {
        const z = quellen[k];
        const r0 = s.moeglich ? rl(s.q) : null;
        z.rl_geb = s.moeglich ? r0.geb - verbraucht("ue1", s.q) : 0;
        z.rl_gub = s.moeglich ? r0.gub - verbraucht("ue2", s.q) - verbraucht("ue3", s.q) : 0;
      });
      const vorher = (key, k) => summe(quellen.slice(0, k).map((z) => z[key]));
      qs.forEach((s, k) => { quellen[k].ue1 = Math.max(Math.min(quellen[k].rl_geb, gebGrenze(k) - vorher("ue1", k)), 0); });
      qs.forEach((s, k) => { quellen[k].ue2 = Math.max(Math.min(quellen[k].rl_gub, akGub - vorher("ue2", k)), 0); });
      const alleUe1 = summe(quellen.map((z) => z.ue1));
      qs.forEach((s, k) => {
        const z = quellen[k];
        z.ue3 = Math.max(Math.min(z.rl_gub - z.ue2, gebGrenze(k) - alleUe1 - vorher("ue3", k)), 0);
      });
      qs.forEach((s, k) => {
        const z = quellen[k];
        if (!gleich(f.erloes_fin, "ja") || !s.erloesMoeglich) { z.erloes = 0; return; }
        const netto = N(verkaeufe[rlIndex(s.q)].berechnet.nettoerloes);
        z.erloes = Math.max(Math.min(netto - verbraucht("erloes", s.q), ak - kreditFest - vorher("erloes", k)), 0);
      });
      b.quellen = quellen;
      x.quellenInfo = qs;
      for (const key of ["rl_geb", "rl_gub", "ue1", "ue2", "ue3"]) b[key] = summe(quellen.map((z) => z[key]));
      b.ue_gesamt = b.ue1 + b.ue2 + b.ue3;
      b.ue_folge = summe(qs.map((s, k) => (s.vorjahr ? quellen[k].ue1 + quellen[k].ue2 + quellen[k].ue3 : 0)));
      b.ue_folge_geb = summe(qs.map((s, k) => (s.vorjahr ? quellen[k].ue1 + quellen[k].ue3 : 0)));
      b.afa_basis = akGeb - b.ue1 - b.ue3;
      b.ak_gub = akGub - b.ue2;
      b.qng_basis = qngBasis(b.afa_basis);
      b.qng_basis_kauf = qngBasis(b.afa_basis + b.ue_folge_geb);
      b.mit_quelle = qs.some((s) => !leer(s.q) && !s.aus) ? 1 : 0;
      b.ak_gesamt = ak;
      b.erloes = summe(quellen.map((z) => z.erloes));
      b.bedarf = gueltig ? ak - b.erloes : 0;
      b.kredit = gueltig && gleich(f.fin_art, FIN_KREDIT)
        ? (!leer(f.kredit_betrag) || !leer(f.kredit_quote) ? kreditFest : Math.max(b.bedarf, 0)) : 0;
      b.kredit2 = b.kredit > 0 ? (!leer(f.kredit2_betrag) ? Math.min(N(f.kredit2_betrag), b.kredit)
        : b.kredit * N(f.kredit2_anteil)) : 0;
      b.eigen = b.bedarf - b.kredit;
      b.miete_qm_jahr = N(f.flaeche) * N(f.miete_qm) > 0 ? f.flaeche * f.miete_qm * 12 : "";
      b.mietsteig_v = leer(f.mietsteig) ? P.NeuMietsteig : N(f.mietsteig);
      const mieteKst = neukaufWert(neukaufFort, f.kst, "miete", N(kj) + 1);
      b.miete_kauf = istZahl(mieteKst) ? mieteKst : istZahl(b.miete_qm_jahr) ? b.miete_qm_jahr
        : N(f.kaufpreis) * N(f.mietrendite);
      b.rendite = leer(b.miete_kauf) || N(f.kaufpreis) <= 0 ? "" : b.miete_kauf / f.kaufpreis;
      b.faktor = N(b.miete_kauf) <= 0 || N(f.kaufpreis) <= 0 ? "" : f.kaufpreis / b.miete_kauf;
      const quellSumme = (fn) => summe(QUELLEN.map((g) => {
        const i = rlIndex(f[g]);
        return i >= 0 ? N(fn(verkaeufe[i])) : 0;
      }));
      const qPreis = quellSumme((v) => v.felder.preis), qMiete = quellSumme((v) => v.berechnet.miete);
      b.rendite_quelle = qPreis > 0 ? qMiete / qPreis : "";
      const qMieteOE = quellSumme((v) => v.berechnet.miete_oe);
      b.erh_quote_quelle = qMieteOE > 0 ? quellSumme((v) => v.berechnet.erh_oe) / qMieteOE : "";
      // Status
      const hinweise = [];
      qs.forEach((s, k) => {
        if (leer(s.q)) return;
        const vor = k ? `; Quelle ${k + 1}: ` : "; ";
        const r0 = rl(s.q);
        let t = "";
        if (k && QUELLEN.slice(0, k).some((g) => gleich(f[g], s.q))) t = "doppelt, keine Übertragung";
        else if (s.aus) t = "Verkauf ausgeschaltet, Rücklage entfällt, keine Übertragung";
        else if (!r0) t = "Rücklage unbekannt, keine Übertragung";
        else if (N(kj) < N(r0.jahr) - VORJAHR) t = "Kauf vor Bildung der Rücklage, keine Übertragung";
        else if (N(kj) > N(r0.fristjahr)) t = "Kauf nach Fristjahr, keine Übertragung";
        if (t) hinweise.push(vor + t);
      });
      hinweise.push(hinweisKredit(f, b));
      if (gleich(f.qng, "ja") && ueberObergrenze) hinweise.push("; § 7b: AK Gebäude je m² über der Baukostenobergrenze, keine Sonder-AfA");
      const h = hinweise.join("");
      x.status = gleich(f.aktiv, "nein") ? STATUS_AUS
        : !pflicht ? "Pflichtfeld fehlt"
        : !aufteilung ? "Anteil G+B oder AK G+B fehlt"
        : neuIds.filter((id) => gleich(id, f.neu_id)).length > 1 ? "NeuID doppelt"
        : objIds.some((id) => gleich(id, f.neu_id)) ? "NeuID wie Bestandsobjekt"
        : N(kj) < P.Startjahr || N(kj) > P.Endjahr ? "Kaufjahr außerhalb Raster"
        : h === "" ? STATUS_OK : h.slice(2, 502);
    });
  }

  function hinweisKredit(f, b) {
    const l = (k) => leer(f[k]);
    let h;
    if (!gleich(f.fin_art, FIN_KREDIT)) {
      h = ["kredit_betrag", "kredit_quote", "kredit_zins", "kredit2_anteil", "kredit2_betrag", "kredit2_zins"].every(l)
        ? "" : "; Kreditangaben ohne Finanzierung Rest = Kredit, kein Kredit";
    } else if (l("kredit_zins")) h = "; Kredit: Zinssatz fehlt";
    else if (gleich(f.tilgungsart, TILGUNG_ENDFAELLIG) && l("laufzeit")) h = "; Kredit: Laufzeit fehlt (endfällig)";
    else if (!gleich(f.tilgungsart, TILGUNG_ENDFAELLIG) && l("tilgung") && l("laufzeit")) h = "; Kredit: Tilgung fehlt";
    else h = "";
    const art2 = !l("kredit2_tilgungsart") ? f.kredit2_tilgungsart : f.tilgungsart;
    const ohneLz = l("kredit2_laufzeit") && l("laufzeit");
    const ohneTilgung = l("kredit2_tilgung") && l("tilgung") && ohneLz;
    let h2;
    if (!gleich(f.fin_art, FIN_KREDIT) || (l("kredit2_anteil") && l("kredit2_betrag"))) {
      h2 = !gleich(f.fin_art, FIN_KREDIT) || l("kredit2_zins") ? "" : "; Zinssatz Zweitkredit ohne Anteil oder Betrag, kein Zweitkredit";
    } else if (l("kredit2_zins")) h2 = "; Zweitkredit: Zinssatz fehlt";
    else if (gleich(art2, TILGUNG_ENDFAELLIG) && ohneLz) h2 = "; Zweitkredit: Laufzeit fehlt (endfällig)";
    else if (!gleich(art2, TILGUNG_ENDFAELLIG) && ohneTilgung) h2 = "; Zweitkredit: Tilgung fehlt";
    else if (!l("kredit2_betrag") && N(f.kredit2_betrag) > b.kredit + 0.005) h2 = "; Zweitkredit: Betrag über dem Kredit, gekürzt";
    else if (gleich(f.kredit2_uebernahme, UEBERNAHME_HAUPT) && ohneLz) h2 = "; Zweitkredit: Übernahme ohne Laufzeit, keine Übernahme";
    else h2 = "";
    return h + h2;
  }

  // ---- Rücklagen: Übertragung, Auflösung, Spiegel ----
  function ruecklagenAbschliessen(ruecklagen, neu, P) {
    ruecklagen.forEach((r) => {
      if (!r) return;
      r.ueb_geb = 0;
      r.ueb_gub = 0;
      neu.forEach((x) => {
        if (!x.berechnet) return;
        QUELLEN.forEach((g, k) => {
          if (!gleich(x.felder[g], r.id)) return;
          const z = x.berechnet.quellen[k];
          r.ueb_geb += N(z.ue1);
          r.ueb_gub += N(z.ue2) + N(z.ue3);
        });
      });
      r.aufloesung = r.betrag - r.ueb_geb - r.ueb_gub;
      r.zuschlag = r.aufloesung * P["6bZuschlag"] * (r.fristjahr - N(r.jahr));
      r.hinweis = r.fristjahr > P.Endjahr && r.aufloesung > 0 ? "Frist endet nach Prognoseende" : "";
    });
  }

  function spiegelRechnen(verkaeufe, ruecklagen, neu, jahre) {
    const s = { jahr: [], gewinne: [], bildung: [], uebertragung: [], aufloesung: [], zuschlag: [],
      steuerwirksam: [], bestand: [] };
    const neuSumme = (key, jahr) => summe(neu.filter((x) => x.berechnet && N(x.felder.kaufjahr) === jahr
      && istZahl(x.felder.kaufjahr)).map((x) => x.berechnet[key]));
    jahre.forEach((t, i) => {
      s.jahr[i] = t;
      s.gewinne[i] = summe(verkaeufe.filter((v) => v.gueltig && N(v.felder.jahr) === t).map((v) => v.berechnet.gewinn));
      s.bildung[i] = summe(ruecklagen.filter((r) => r && N(r.jahr) === t).map((r) => r.betrag));
      s.uebertragung[i] = neuSumme("ue_gesamt", t) - neuSumme("ue_folge", t) + neuSumme("ue_folge", t - 1);
      s.aufloesung[i] = summe(ruecklagen.filter((r) => r && r.fristjahr === t).map((r) => r.aufloesung));
      s.zuschlag[i] = summe(ruecklagen.filter((r) => r && r.fristjahr === t).map((r) => r.zuschlag));
      s.steuerwirksam[i] = s.gewinne[i] - s.bildung[i] + s.aufloesung[i] + s.zuschlag[i];
      s.bestand[i] = (i === 0 ? 0 : s.bestand[i - 1]) + s.bildung[i] - s.uebertragung[i] - s.aufloesung[i];
    });
    return s;
  }

  // ---- Darlehen der Neuobjekte ----
  function darlehenRechnen(neu, P, jahre) {
    return neu.map((x) => {
      if (!x.berechnet) return null;
      const f = x.felder, b = x.berechnet;
      const zweit = kreditPlan(f, b, P, jahre, true, null);
      const haupt = kreditPlan(f, b, P, jahre, false, zweit);
      return { haupt, zweit };
    });
  }

  function kreditPlan(f, b, P, jahre, zweit, partner) {
    const eigen = (k, k2) => (!zweit ? f[k] : !leer(f[k2]) ? f[k2] : f[k]);
    const betrag = zweit ? N(b.kredit2) : N(b.kredit) - N(b.kredit2);
    const zins = zweit ? N(f.kredit2_zins) : N(f.kredit_zins);
    const partnerUe = partner ? partner.ue_betrag : "";
    const da = zweit ? betrag > 0.005 : betrag > 0.005 || N(partnerUe) > 0.005;
    if (!da) return null;
    const d = { id: f.neu_id, teil: zweit ? "Zweitkredit" : "Hauptkredit", kaufjahr: f.kaufjahr, betrag,
      zinssatz: zins };
    const artEin = eigen("tilgungsart", "kredit2_tilgungsart");
    d.art = leer(artEin) ? TILGUNG_ANNUITAET : artEin;
    d.tilgungssatz = N(eigen("tilgung", "kredit2_tilgung"));
    const lz = eigen("laufzeit", "kredit2_laufzeit");
    d.laufzeit = N(lz) > 0 ? N(lz) : "";
    d.tilgungsfrei = N(eigen("tilgungsfrei", "kredit2_tilgungsfrei"));
    const endf = gleich(d.art, TILGUNG_ENDFAELLIG), linear = gleich(d.art, TILGUNG_LINEAR);
    d.rate = endf ? 0 : linear ? betrag * d.tilgungssatz : betrag * (zins + d.tilgungssatz);
    const kj = N(d.kaufjahr);
    if (zweit) {
      // leer = auf Hauptkredit, wenn der länger läuft (keine oder die längere Laufzeit)
      const lzHaupt = N(f.laufzeit);
      const ueber = gleich(f.kredit2_uebernahme, UEBERNAHME_HAUPT)
        || (leer(f.kredit2_uebernahme) && (lzHaupt <= 0 || lzHaupt > N(d.laufzeit)));
      d.ue_jahr = ueber && istZahl(d.laufzeit) ? kj + d.laufzeit : "";
      d.rate_neu = "";
    } else {
      d.ue_jahr = istZahl(partnerUe) ? partner.ue_jahr : "";
      d.ue_betrag = istZahl(partnerUe) ? partnerUe : "";
      const faktor = endf ? 0 : linear ? d.tilgungssatz : d.zinssatz + d.tilgungssatz;
      d.rate_neu = istZahl(d.ue_betrag) ? d.rate + d.ue_betrag * faktor : "";
    }
    d.zins = [];
    d.tilgung = [];
    d.restschuld = [];
    jahre.forEach((t, i) => {
      const vor = i === 0 ? 0 : N(d.restschuld[i - 1]);
      d.zins[i] = vor * zins;
      const rate = istZahl(d.rate_neu) && t > N(d.ue_jahr) ? d.rate_neu : d.rate;
      const regulaer = endf ? 0 : linear ? rate : Math.max(rate - d.zins[i], 0);
      const ende = istZahl(d.laufzeit) && t >= kj + d.laufzeit && (!zweit || !istZahl(d.ue_jahr));
      d.tilgung[i] = t <= kj ? 0 : Math.min(vor, ende ? vor : t <= kj + d.tilgungsfrei ? 0 : regulaer);
      const rest = zweit ? (t === N(d.ue_jahr) ? 0 : vor - d.tilgung[i])
        : vor - d.tilgung[i] + (t === N(d.ue_jahr) ? N(d.ue_betrag) : 0);
      d.restschuld[i] = t < kj ? 0 : t === kj ? betrag : rest;
    });
    if (zweit) {
      d.ue_betrag = istZahl(d.ue_jahr) && d.ue_jahr <= P.Endjahr
        ? Math.max(betrag - summe(jahre.map((t, i) => (t <= d.ue_jahr ? d.tilgung[i] : 0))), 0) : "";
    }
    const letzte = d.restschuld[jahre.length - 1];
    if (letzte > 0.005) d.getilgt = "nach " + P.Endjahr;
    else {
      const mit = jahre.filter((t, i) => d.restschuld[i] > 0.005);
      d.getilgt = (mit.length ? Math.max(...mit) : 0) + 1;
    }
    return d;
  }

  // ---- Neukauf-Kostenstellen: fortgeschriebene Planwerte ----
  function neukaufFortschreiben(eingabe, P, neu) {
    const n = P.Prognosejahre + 1;   // Basisjahr bis Endjahr
    const fort = {};
    for (const nk of eingabe.neukauf || []) {
      if (leer(nk.kst)) continue;
      const nr = neu.findIndex((x) => gleich(x.felder.kst, nk.kst));
      const x = nr >= 0 ? neu[nr] : null;
      const mitId = x && !leer(x.felder.neu_id);
      const qm = mitId && N(x.felder.flaeche) * N(x.felder.miete_qm) > 0 ? x.felder.flaeche * x.felder.miete_qm * 12 : "";
      const kj = x && istZahl(x.felder.kaufjahr) ? x.felder.kaufjahr : "";
      const msteig = mitId ? (leer(x.felder.mietsteig) ? P.NeuMietsteig : N(x.felder.mietsteig)) : P.NeuMietsteig;
      const warte = x ? N(x.felder.miet_warte) : 0;
      const bis = x && istZahl(x.felder.miet_bis) ? x.felder.miet_bis : 10000;
      const werte = {};
      for (const pos of ["miete", "einnahmen", "erhaltung", "ausgaben"]) {
        const satz = pos === "miete" || pos === "einnahmen" ? msteig
          : pos === "erhaltung" ? P.Erhaltsteig : P.Kostensteig;
        const ein = (nk.werte || {})[pos] || {};
        const reihe = [];
        for (let i = 0; i < n; i++) {
          const jahr = P.Basisjahr + i;
          const wert = ein[jahr] !== undefined ? ein[jahr] : ein[String(jahr)];
          let annahme = "";
          if (pos === "miete") {
            annahme = istZahl(qm) && jahr > N(kj)
              ? qm * Math.pow(1 + satz, Math.max(Math.min(jahr - N(kj), bis) - warte, 0)) : "";
          } else if (pos === "erhaltung" || pos === "ausgaben") {
            const m = werte.miete[i];
            annahme = istZahl(qm) && istZahl(m) ? m * (pos === "erhaltung" ? P.QmErhQuote : P.QmVerwQuote) : "";
          }
          const vorher = reihe.length > 0 && Object.keys(ein).some((j) => Number(j) >= P.Basisjahr
            && Number(j) < jahr && istZahl(ein[j]) && ein[j] !== 0);
          if (i === 0) reihe.push(N(wert) !== 0 ? wert : annahme);
          else if (istZahl(wert) && (wert !== 0 || vorher)) reihe.push(wert);
          else if (vorher) {
            const steig = pos === "miete" && jahr > N(kj) && (jahr <= N(kj) + warte || jahr > N(kj) + bis) ? 0 : satz;
            reihe.push(N(reihe[i - 1]) * (1 + steig));
          } else reihe.push(annahme);
        }
        werte[pos] = reihe;
      }
      fort[String(nk.kst).toLowerCase()] = werte;
    }
    return fort;
  }

  function neukaufWert(fort, kst, pos, jahr) {
    if (leer(kst)) return "";
    const w = fort[String(kst).toLowerCase()];
    if (!w) return "";
    const i = jahr - fort.__basisjahr;
    const v = w[pos][i];
    return v === undefined ? "" : v;
  }

  // ---- Prognose der Neuobjekte ----
  function prognoseNeu(x, P, jahre, darlehen, fort) {
    const f = x.felder, b = x.berechnet;
    const kj = N(f.kaufjahr), gueltig = b.gueltig === 1;
    const s = {};
    const spalten = ["id", "jahr", "aktiv", "miete", "einnahmen", "erhaltung", "ausgaben", "afa", "afa_qng",
      "buchwert", "ergebnis", "verkehrswert", "bestand", "neu", "buchwert_gub", "buchwert_halten",
      "mit_quelle", "afa_ohne6b", "afa_qng_ohne6b", "buchwert_ohne6b", "buchwert_gub_ohne6b",
      "erhaltung_halten", "miete_halten", "einnahmen_halten", "quelle", "afa_halten", "kredit_zins",
      "tilgung", "abloesung", "restschuld", "zins_halten", "tilgung_halten", "restschuld_halten",
      "ergebnis_halten", "wirkung", "verkauf_nr", "unternehmen"];
    spalten.forEach((k) => (s[k] = new Array(jahre.length)));
    const qm = b.miete_qm_jahr, mitQm = istZahl(qm);
    const mietBis = istZahl(f.miet_bis) ? f.miet_bis : null;
    const afaRechnen = (basis, folge, qb, qbk, bw, afaK, qngK, i, t, aktiv) => {
      const bwVor = i === 0 ? 0 : s[bw][i - 1];
      const bwStart = t === kj ? basis + folge : t === kj + 1 ? Math.max(bwVor - folge, 0) : bwVor;
      const qng = t >= kj && t < kj + P.QNGJahre ? Math.min((t === kj ? qbk : qb) * P.QNGSatz, bwStart) : 0;
      const rest = bwStart - qng;
      const satz = N(f.afa_satz);
      let regulaer = 0;
      if (aktiv) {
        if (gleich(f.afa_methode, "degressiv")) {
          const deg = bwStart * P.AfADegressiv;
          const wahl = gleich(P.AfAWechsel, "ja") ? Math.max(deg, bwStart / Math.max(1 / satz - (t - kj - 1), 1)) : deg;
          regulaer = Math.min(wahl, rest);
        } else {
          const restwert = Math.max(basis - qbk * P.QNGSatz - (P.QNGJahre - 1) * qb * P.QNGSatz
            - (P.QNGJahre - 1) * basis * satz, 0);
          const linBasis = qb > 0 && P.QNGJahre > 0 && t >= kj + P.QNGJahre ? restwert : basis;
          regulaer = Math.min(linBasis * satz, rest);
        }
      }
      s[qngK][i] = qng;
      s[afaK][i] = qng + regulaer;
      s[bw][i] = gueltig ? (t < kj ? 0 : Math.max(bwStart - s[afaK][i], 0)) : 0;
    };
    const kst = (pos, t) => neukaufWert(fort, f.kst, pos, t);
    const dl = (key, i) => {
      if (!gueltig || !darlehen) return 0;
      return N(darlehen.haupt ? darlehen.haupt[key][i] : 0) + N(darlehen.zweit ? darlehen.zweit[key][i] : 0);
    };
    jahre.forEach((t, i) => {
      s.id[i] = f.neu_id;
      s.jahr[i] = t;
      const aktiv = gueltig && t > kj ? 1 : 0;
      s.aktiv[i] = aktiv;
      const mk = kst("miete", t);
      const basisMiete = mitQm ? qm : N(f.kaufpreis) * N(f.mietrendite);
      const jahreSeit = Math.min(t - kj, mietBis === null ? t - kj : mietBis);
      s.miete[i] = aktiv ? (istZahl(mk) ? mk
        : basisMiete * Math.pow(1 + b.mietsteig_v, Math.max(jahreSeit - N(f.miet_warte), 0))) : 0;
      s.einnahmen[i] = aktiv ? N(kst("einnahmen", t)) : 0;
      const ek = kst("erhaltung", t);
      s.erhaltung[i] = aktiv ? (istZahl(ek) ? ek : mitQm ? s.miete[i] * P.QmErhQuote
        : N(f.kaufpreis) * N(f.erhaltungsquote) * (t - kj <= P.NeuErhAnlaufJahre ? P.NeuErhAnlaufFaktor : 1)
          * Math.pow(1 + P.Erhaltsteig, t - kj)) : 0;
      const ak = kst("ausgaben", t);
      s.ausgaben[i] = aktiv ? (istZahl(ak) ? ak : mitQm ? s.miete[i] * P.QmVerwQuote : 0) : 0;
      afaRechnen(b.afa_basis, N(b.ue_folge_geb), b.qng_basis, b.qng_basis_kauf, "buchwert", "afa", "afa_qng", i, t, aktiv);
      afaRechnen(b.ak_geb_neu, 0, b.qng_basis_ohne6b, b.qng_basis_ohne6b, "buchwert_ohne6b", "afa_ohne6b",
        "afa_qng_ohne6b", i, t, aktiv);
      s.ergebnis[i] = (s.miete[i] + s.einnahmen[i] - s.erhaltung[i] - s.ausgaben[i]) * aktiv - s.afa[i];
      s.verkehrswert[i] = gueltig && t >= kj ? N(f.kaufpreis) * Math.pow(1 + P.Wertsteig, t - kj) : 0;
      s.bestand[i] = gueltig && t >= kj ? 1 : 0;
      s.neu[i] = 1;
      s.buchwert_gub[i] = gueltig && t >= kj ? b.ak_gub + (t === kj ? N(b.ue_folge) - N(b.ue_folge_geb) : 0) : 0;
      for (const k of ["buchwert_halten", "afa_halten", "erhaltung_halten", "miete_halten", "einnahmen_halten",
        "zins_halten", "tilgung_halten", "restschuld_halten", "ergebnis_halten", "abloesung", "verkauf_nr",
        "unternehmen"]) s[k][i] = 0;
      s.mit_quelle[i] = b.mit_quelle;
      s.buchwert_gub_ohne6b[i] = gueltig && t >= kj ? b.ak_gub_neu : 0;
      s.quelle[i] = x.status === STATUS_OK && !leer(f.quelle) ? f.quelle : "";
      s.kredit_zins[i] = dl("zins", i);
      s.tilgung[i] = dl("tilgung", i);
      s.restschuld[i] = dl("restschuld", i);
      s.wirkung[i] = s.ergebnis[i] - s.kredit_zins[i];
    });
    return s;
  }

  // ---- Liquidität und Auswertung je Szenario ----
  function szenarienRechnen(ctx) {
    const { P, jahre, objekte, bestand, neuPrg, neu, verkaeufe, spiegel } = ctx;
    const zeilen = [];
    bestand.forEach((s, j) => { if (s) zeilen.push(s); });
    neuPrg.forEach((s) => { if (s) zeilen.push(s); });
    const prgSumme = (k, i, filter) => summe(zeilen.filter((s) => !filter || filter(s, i)).map((s) => s[k][i]));
    const ausschuettung = (t) => {
      const ab = istZahl(P.AusschuettungAb) ? P.AusschuettungAb : P.Startjahr;
      return t >= ab ? N(P.Ausschuettung) * Math.pow(1 + N(P.AusschuettungSteig), t - ab) : 0;
    };
    const objMitId = objekte.filter(Boolean);
    const erg = {};
    for (const sz of SZENARIEN) {
      const L = {};
      ["jahr", "einnahmen", "ausgaben", "afa", "ergebnis", "verkauf", "zins", "kreditzins", "zve",
        "vortrag_genutzt", "bemessung", "vortrag", "steuer", "verkaufserloes", "rueckfluss", "kauf", "kredit",
        "tilgung", "abloesung", "ausschuettung", "zufluss", "kum", "restschuld"].forEach((k) => (L[k] = []));
      const nurB = sz === "B" ? (s, i) => s.mit_quelle[i] === 0 : null;
      const neuImSz = (x) => x.berechnet && x.berechnet.gueltig === 1 && (sz !== "B" || x.berechnet.mit_quelle === 0);
      jahre.forEach((t, i) => {
        L.jahr[i] = t;
        if (sz === "Baseline") {
          L.einnahmen[i] = prgSumme("miete_halten", i) + prgSumme("einnahmen_halten", i);
          L.ausgaben[i] = prgSumme("erhaltung_halten", i)
            + summe(objMitId.map((o) => o.felder.weitere_ausgaben)) * Math.pow(1 + P.Kostensteig, t - P.Basisjahr);
          const bwVor = i === 0 ? summe(objMitId.map((o) => o.felder.restbuchwert)) : prgSumme("buchwert_halten", i - 1);
          L.afa[i] = bwVor - prgSumme("buchwert_halten", i);
          L.verkauf[i] = 0;
          L.verkaufserloes[i] = 0;
          L.rueckfluss[i] = 0;
          L.kauf[i] = 0;
          L.kreditzins[i] = prgSumme("zins_halten", i);
          L.kredit[i] = 0;
          L.tilgung[i] = prgSumme("tilgung_halten", i);
          L.abloesung[i] = 0;
          L.restschuld[i] = prgSumme("restschuld_halten", i);
        } else {
          L.einnahmen[i] = prgSumme("miete", i, nurB) + prgSumme("einnahmen", i, nurB);
          L.ausgaben[i] = prgSumme("erhaltung", i, nurB) + prgSumme("ausgaben", i, nurB);
          L.afa[i] = prgSumme(sz === "A" ? "afa" : "afa_ohne6b", i, nurB);
          L.verkauf[i] = sz === "A" ? spiegel.steuerwirksam[i] : spiegel.gewinne[i];
          const erloese = summe(verkaeufe.filter((v) => v.gueltig && N(v.felder.jahr) === t).map((v) => v.berechnet.nettoerloes));
          L.verkaufserloes[i] = erloese;
          L.rueckfluss[i] = erloese - spiegel.gewinne[i];
          L.kauf[i] = summe(neu.filter((x) => neuImSz(x) && N(x.felder.kaufjahr) === t)
            .map((x) => N(x.felder.kaufpreis) + N(x.felder.nebenkosten)));
          L.kreditzins[i] = prgSumme("kredit_zins", i, nurB);
          L.kredit[i] = summe(neu.filter((x) => neuImSz(x) && N(x.felder.kaufjahr) === t).map((x) => x.berechnet.kredit));
          L.tilgung[i] = prgSumme("tilgung", i, nurB);
          L.abloesung[i] = prgSumme("abloesung", i, nurB);
          L.restschuld[i] = prgSumme("restschuld", i, nurB);
        }
        const kumVor = i === 0 ? 0 : L.kum[i - 1];
        const vortragVor = i === 0 ? 0 : L.vortrag[i - 1];
        L.ergebnis[i] = L.einnahmen[i] - L.ausgaben[i] - L.afa[i];
        L.zins[i] = kumVor * P.Alternativrendite;
        L.zve[i] = L.ergebnis[i] + L.verkauf[i] + L.zins[i] - L.kreditzins[i];
        L.vortrag_genutzt[i] = Math.min(vortragVor, Math.max(L.zve[i], 0));
        L.bemessung[i] = Math.max(L.zve[i], 0) - L.vortrag_genutzt[i];
        L.vortrag[i] = vortragVor - L.vortrag_genutzt[i] + Math.max(-L.zve[i], 0);
        L.steuer[i] = L.bemessung[i] * P.SteuersatzEff + (L.zve[i] - L.bemessung[i]) * P.VerlustSatz;
        L.ausschuettung[i] = ausschuettung(t);
        L.zufluss[i] = L.einnahmen[i] - L.ausgaben[i] + L.zins[i] - L.kreditzins[i] + L.verkaufserloes[i]
          - L.steuer[i] - L.kauf[i] + L.kredit[i] - L.tilgung[i] - L.ausschuettung[i];
        L.kum[i] = kumVor + L.zufluss[i];
      });
      // Auswertung
      const A = {};
      ["jahr", "ergebnis", "verkauf", "zins", "kreditzins", "guv", "steuer", "nach_steuer", "steuer_kum",
        "verkehrswert", "buchwert", "stille_reserven", "ruecklage", "vortrag", "liquiditaet", "restschuld",
        "vermoegen", "latente_steuer", "vermoegen_netto"].forEach((k) => (A[k] = []));
      const imBestand = sz === "Baseline" ? (s, i) => s.neu[i] === 0
        : (s, i) => s.bestand[i] === 1 && (sz !== "B" || s.mit_quelle[i] === 0);
      const geb = sz === "Baseline" ? "buchwert_halten" : sz === "A" ? "buchwert" : "buchwert_ohne6b";
      const gub = sz === "A" || sz === "Baseline" ? "buchwert_gub" : "buchwert_gub_ohne6b";
      jahre.forEach((t, i) => {
        A.jahr[i] = t;
        A.ergebnis[i] = L.ergebnis[i];
        A.verkauf[i] = L.verkauf[i];
        A.zins[i] = L.zins[i];
        A.kreditzins[i] = L.kreditzins[i];
        A.guv[i] = A.ergebnis[i] + A.verkauf[i] + A.zins[i] - A.kreditzins[i];
        A.steuer[i] = L.steuer[i];
        A.nach_steuer[i] = A.guv[i] - A.steuer[i];
        A.steuer_kum[i] = (i === 0 ? 0 : A.steuer_kum[i - 1]) + A.steuer[i];
        A.verkehrswert[i] = prgSumme("verkehrswert", i, imBestand);
        A.buchwert[i] = prgSumme(geb, i, imBestand) + prgSumme(gub, i, imBestand);
        A.stille_reserven[i] = A.verkehrswert[i] - A.buchwert[i];
        A.ruecklage[i] = sz === "A" ? spiegel.bestand[i] : 0;
        A.vortrag[i] = L.vortrag[i];
        A.liquiditaet[i] = L.kum[i];
        A.restschuld[i] = L.restschuld[i];
        A.vermoegen[i] = A.verkehrswert[i] + A.liquiditaet[i] - A.restschuld[i];
        A.latente_steuer[i] = latenteSteuer(A.stille_reserven[i] + A.ruecklage[i], A.vortrag[i], P);
        A.vermoegen_netto[i] = A.vermoegen[i] - A.latente_steuer[i];
      });
      erg[sz] = { liq: L, aus: A };
    }
    return erg;
  }

  function latenteSteuer(betrag, vortrag, P) {
    const nachVortrag = Math.max(betrag - vortrag, 0);
    return nachVortrag * P.SteuersatzEff + (betrag - nachVortrag) * P.VerlustSatz;
  }

  // ---- Vergleich (Kennzahlen am Ende bzw. Summe über alle Jahre) ----
  const VERGLEICH = [
    ["Endvermögen nach latenter Steuer", "aus", "vermoegen_netto", "ende"],
    ["Verkehrswert Immobilien", "aus", "verkehrswert", "ende"],
    ["Liquidität (Alternativanlage)", "aus", "liquiditaet", "ende"],
    ["latente Steuer", "aus", "latente_steuer", "ende"],
    ["Buchwert Immobilien", "aus", "buchwert", "ende"],
    ["stille Reserven", "aus", "stille_reserven", "ende"],
    ["§ 6b-Rücklage", "aus", "ruecklage", "ende"],
    ["Verlustvortrag", "aus", "vortrag", "ende"],
    ["Mieten und weitere Einnahmen", "liq", "einnahmen", "summe"],
    ["Erhaltung und weitere Ausgaben", "liq", "ausgaben", "summe"],
    ["AfA Gebäude", "liq", "afa", "summe"],
    ["laufendes Ergebnis", "liq", "ergebnis", "summe"],
    ["steuerwirksam aus Verkauf und Rücklage", "liq", "verkauf", "summe"],
    ["Zinsertrag Alternativanlage", "liq", "zins", "summe"],
    ["Steuer", "liq", "steuer", "summe"],
    ["Verkaufserlöse netto", "liq", "verkaufserloes", "summe"],
    ["Kauf Neuobjekte", "liq", "kauf", "summe"],
    ["Restschuld Darlehen", "aus", "restschuld", "ende"],
    ["Zinsen Darlehen", "liq", "kreditzins", "summe"],
    ["Ausschüttungen", "liq", "ausschuettung", "summe"],
  ];

  function vergleichRechnen(sz) {
    return VERGLEICH.map(([kennzahl, tab, key, art]) => {
      const z = { kennzahl, tabelle: tab, key, art };
      for (const s of SZENARIEN) {
        const reihe = sz[s][tab][key];
        z[s] = art === "ende" ? reihe[reihe.length - 1] : summe(reihe);
      }
      z.diffB = z.A - z.B;
      z.diffC = z.A - z.C;
      z.diffBaseline = z.A - z.Baseline;
      return z;
    });
  }

  // ---- Plausibilitätsprüfungen ----
  function pruefungRechnen(ctx) {
    const { P, objekte, verkaeufe, ruecklagen, neu, anlagen, darlehen, sz, afaPlan, eingabe } = ctx;
    const anzahl = {};
    const nichtOk = (liste, ...ausnahmen) => liste.filter((s) => !leer(s) && s !== STATUS_OK && !ausnahmen.includes(s)).length;
    anzahl.objekte = nichtOk(objekte.filter(Boolean).map((o) => o.status));
    anzahl.verkaeufe = nichtOk(verkaeufe.map((v) => v.status), STATUS_6B_UNZULAESSIG, STATUS_AUS);
    anzahl.neuobjekte = nichtOk(neu.map((x) => x.status), STATUS_AUS);
    anzahl.steuerwelt = [STEUERWELT_GMBH, STEUERWELT_KG].some((w) => gleich(P.Steuerwelt, w)) ? 0 : 1;
    anzahl.kg_vereinfacht = gleich(P.Steuerwelt, STEUERWELT_KG) ? 1 : 0;
    anzahl.vorbesitz = verkaeufe.filter((v) => v.status === STATUS_6B_UNZULAESSIG).length;
    const genannt = (r) => neu.filter((x) => QUELLEN.some((g) => gleich(x.felder[g], r.id))).length;
    const rest = ruecklagen.filter((r) => r && r.aufloesung > 0.005);
    anzahl.teiluebertrag = rest.filter((r) => genannt(r) > 0).length;
    anzahl.ohne_reinvest = rest.filter((r) => genannt(r) === 0).length;
    const gueltige = verkaeufe.filter((v) => v.gueltig);
    anzahl.grundstueckshandel = gueltige.filter((v) => gueltige.filter((w) =>
      N(w.felder.jahr) > N(v.felder.jahr) - P.DOJahre && N(w.felder.jahr) <= N(v.felder.jahr)).length > P.DOGrenze).length;
    anzahl.frist_ende = ruecklagen.filter((r) => r && r.hinweis !== "").length;
    anzahl.liquiditaet = sz.A.liq.kum.filter((k) => k < N(P.MindestLiq) - 0.005).length;
    const verkauftAnz = (id) => verkaeufe.filter((v) => v.aktiv && imPaket(v.umfang, id)).length;
    anzahl.kritisch = objekte.filter((o) => o && verkauftAnz(o.felder.objekt_id) > 0 && o.kritischOhneVerkauf > 0).length
      + verkaeufe.filter((v) => !leer(v.felder.objekt_id) && v.aktiv && v.herkunft.preis === "annahme").length;
    anzahl.bwa_zuordnung = 0;   // BWA-Blätter gibt es im Browser nicht
    anzahl.annahmen = objekte.filter((o) => o && o.annahmen > 0).length;
    anzahl.anlagen = anlagen.filter((a) => a && a.status !== STATUS_OK && a.status !== ANL_NICHT_IM_MODELL).length;
    anzahl.anlagen_bez = anlagen.filter((a) => a && /^(bezeichnung|mehrdeutig)/i.test(text(a.felder.zuordnung))).length;
    anzahl.baujahr = summe(objekte.filter(Boolean).map((o) => o.anl.anl_bau_offen));
    anzahl.kst_abweichung = 0;   // keine Kostenstellenblätter im Browser
    const nkIds = (eingabe.neukauf || []).map((n) => n.kst);
    anzahl.neukauf_kst = neu.filter((x) => !leer(x.felder.kst) && !nkIds.some((k) => gleich(k, x.felder.kst))).length;
    anzahl.ak_aufteilung = neu.filter((x) => x.berechnet && Math.abs(x.berechnet.ak_abweichung) > 0.5).length;
    anzahl.vorjahr = neu.filter((x) => x.berechnet && x.berechnet.ue_folge > 0.005).length;
    anzahl.verkauf_unternehmen = objekte.filter((o) => o && gleich(o.felder.art, ART_UNTERNEHMEN)
      && verkaufsjahr(verkaeufe, o.felder.objekt_id).jahr !== 9999).length;
    anzahl.bruecke = 0;
    anzahl.cashflow = 0;
    const L = sz.A.liq;
    anzahl.kapitaldienst = L.jahr.filter((t, i) => {
      const kd = L.kreditzins[i] + L.tilgung[i] - L.abloesung[i];
      return kd > 0.005 && L.einnahmen[i] - L.ausgaben[i] < kd;
    }).length;
    anzahl.afa_plan = objekte.filter((o) => o && Object.values((afaPlan || {})[o.felder.objekt_id] || {}).some(istZahl)).length;
    anzahl.anlagen_afa = objekte.filter((o) => o && istZahl(o.anl.anl_diff) && Math.abs(o.anl.anl_diff) > 1).length;
    anzahl.restschuld = darlehen.filter(Boolean).reduce((a, d) => a
      + [d.haupt, d.zweit].filter((k) => k && /^nach/.test(String(k.getilgt))).length, 0);
    const zinsGrob = objekte.filter((o) => o && N(o.felder.zinsen) > 0 && !(N(o.felder.restschuld) > 0));
    anzahl.zins_verkauf = zinsGrob.filter((o) => verkauftAnz(o.felder.objekt_id) > 0).length;
    anzahl.zins_grob = zinsGrob.length;
    return anzahl;
  }

  // ---- Hauptfunktion ----
  function rechne(eingabe) {
    const P = parameterBerechnen(eingabe.parameter || {});
    const jahre = [];
    for (let t = P.Startjahr; t <= P.Endjahr; t++) jahre.push(t);
    const objektIds = (eingabe.objekte || []).map((o) => o.objekt_id).filter((id) => !leer(id));
    const anlagen = anlagenRechnen(eingabe, P, objektIds, jahre);
    const objekte = objekteRechnen(eingabe, P, anlagen);
    const verkaeufe = verkaeufeVorbereiten(eingabe, P, objekte);
    const bestand = objekte.map((o) => (o ? prognoseBestand(o, P, jahre, anlagen, verkaeufe, eingabe.afa_plan) : null));
    verkaeufeRechnen(verkaeufe, objekte, bestand, P, jahre);
    const ruecklagen = ruecklagenRechnen(verkaeufe, P);
    const neu = neuobjekteEingaben(eingabe, P, verkaeufe);
    // Neukauf-KSt: die Annahmen nach m² brauchen Fläche, Kaufjahr und Mietsteigerung der
    // Neuobjekte, aber nichts aus deren Rechnung
    const fort = neukaufFortschreiben(eingabe, P, neu);
    fort.__basisjahr = P.Basisjahr;
    neuobjekteRechnen(neu, ruecklagen, verkaeufe, objekte, P, fort);
    ruecklagenAbschliessen(ruecklagen, neu, P);
    const darlehen = darlehenRechnen(neu, P, jahre);
    const neuPrg = neu.map((x, r) => (x.berechnet ? prognoseNeu(x, P, jahre, darlehen[r], fort) : null));
    // Vergleichswerte der Neuobjekte aus ihrer Prognose
    neu.forEach((x, r) => {
      if (!x.berechnet) return;
      const b = x.berechnet, kj = x.felder.kaufjahr;
      const durch = (k) => {
        if (b.gueltig !== 1 || leer(kj)) return "";
        const bisJ = Math.min(N(kj) + VERGLEICH_JAHRE, P.Endjahr);
        if (bisJ <= N(kj)) return "";
        return summe(jahre.map((t, i) => (t > N(kj) && t <= bisJ ? neuPrg[r][k][i] : 0))) / (bisJ - N(kj));
      };
      b.miete_oe = durch("miete");
      b.erh_oe = durch("erhaltung");
      b.erh_quote = leer(b.erh_oe) || N(b.miete_oe) <= 0 ? "" : b.erh_oe / b.miete_oe;
    });
    const spiegel = spiegelRechnen(verkaeufe, ruecklagen, neu, jahre);
    const sz = szenarienRechnen({ P, jahre, objekte, bestand, neuPrg, neu, verkaeufe, spiegel });
    const vergleich = vergleichRechnen(sz);
    const pruefung = pruefungRechnen({ P, objekte, verkaeufe, ruecklagen, neu, anlagen, darlehen, sz,
      afaPlan: eingabe.afa_plan, eingabe });
    return { parameter: P, jahre, anlagen, objekte, verkaeufe, ruecklagen, neuobjekte: neu, darlehen,
      prognose: { bestand, neu: neuPrg }, spiegel, szenarien: sz, vergleich, pruefung,
      neukauf: fort };
  }

  return { rechne, latenteSteuer, konstanten: { STATUS_OK, STATUS_AUS, STATUS_6B_UNZULAESSIG, SZENARIEN, QUELLEN } };
});
