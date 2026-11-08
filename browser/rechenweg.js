/*
 * Rechenweg je Zahl: woraus sich ein Wert zusammensetzt, mit den Werten der Teile.
 *
 * Jede Zahl der Oberfläche hat einen Pfad, z. B. "prg:b:0:buchwert:3" (Prognose,
 * Bestandsobjekt 1, Buchwert, viertes Prognosejahr). erklaere(pfad, R, M) liefert Titel,
 * Formel in Worten und die Teile mit ihren Pfaden, so dass man sich durch die Rechnung
 * klicken kann. Die Formeln beschreiben die Rechnung in rechenwerk.js.
 */
(function (wurzel, fabrik) {
  if (typeof module === "object" && module.exports) module.exports = fabrik();
  else wurzel.Rechenweg = fabrik();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const leer = (v) => v === null || v === undefined || v === "";
  const N = (v) => (typeof v === "number" && isFinite(v) ? v : 0);

  /** Wert hinter einem Pfad. */
  function wertVon(pfad, R) {
    const t = pfad.split(":");
    const z = (x) => Number(x);
    switch (t[0]) {
      case "par": return R.parameter[t[1]];
      case "obj": return R.objekte[z(t[1])] ? R.objekte[z(t[1])].felder[t[2]] : "";
      case "prg": {
        const s = (t[1] === "b" ? R.prognose.bestand : R.prognose.neu)[z(t[2])];
        return s ? s[t[3]][z(t[4])] : "";
      }
      case "vk": {
        const v = R.verkaeufe[z(t[1])];
        if (!v) return "";
        if (t[2] === "status") return v.status;
        return v.berechnet && t[2] in v.berechnet ? v.berechnet[t[2]] : v.felder[t[2]];
      }
      case "rl": return R.ruecklagen[z(t[1])] ? R.ruecklagen[z(t[1])][t[2]] : "";
      case "rls": return R.spiegel[t[1]][z(t[2])];
      case "ne": {
        const x = R.neuobjekte[z(t[1])];
        if (!x) return "";
        if (t[2] === "status") return x.status;
        return x.berechnet && t[2] in x.berechnet ? x.berechnet[t[2]] : x.felder[t[2]];
      }
      case "neq": {
        const x = R.neuobjekte[z(t[1])];
        return x && x.berechnet ? x.berechnet.quellen[z(t[2])][t[3]] : "";
      }
      case "dl": {
        const d = R.darlehen[z(t[1])];
        const k = d && d[t[2]];
        if (!k) return "";
        return t.length > 4 ? k[t[3]][z(t[4])] : k[t[3]];
      }
      case "liq": return R.szenarien[t[1]].liq[t[2]][z(t[3])];
      case "aus": return R.szenarien[t[1]].aus[t[2]][z(t[3])];
      case "vg": return R.vergleich[z(t[1])][t[2]];
      case "anl": {
        const a = R.anlagen[z(t[1])];
        if (!a) return "";
        return t.length > 3 ? a.jahre[z(t[3])] : t[2] in a ? a[t[2]] : a.felder[t[2]];
      }
      default: return "";
    }
  }

  /** Name des Objekts bzw. Vorgangs einer Zeile, für Beschriftungen. */
  function wer(pfad, R) {
    const t = pfad.split(":");
    if (t[0] === "prg") {
      const s = (t[1] === "b" ? R.prognose.bestand : R.prognose.neu)[Number(t[2])];
      return s ? s.id[0] : "";
    }
    if (t[0] === "obj" && R.objekte[Number(t[1])]) return R.objekte[Number(t[1])].felder.objekt_id;
    if (t[0] === "vk" && R.verkaeufe[Number(t[1])]) return "Verkauf " + (R.verkaeufe[Number(t[1])].felder.objekt_id || "");
    if ((t[0] === "ne" || t[0] === "neq" || t[0] === "dl") && R.neuobjekte[Number(t[1])])
      return R.neuobjekte[Number(t[1])].felder.neu_id;
    if (t[0] === "rl" && R.ruecklagen[Number(t[1])]) return R.ruecklagen[Number(t[1])].id;
    return "";
  }

  // Bezeichnungen der Parameter und Felder kommen aus den Metadaten (modelle.py)
  function bezeichnung(pfad, M) {
    const t = pfad.split(":");
    const such = (liste, key) => (liste || []).find((x) => x.key === key || x.name === key);
    switch (t[0]) {
      case "par": { const p = such(M.parameter, t[1]); return p ? p.bezeichnung : t[1]; }
      case "obj": { const f = such(M.felder.objekte, t[2]); return f ? f.ueberschrift : t[2]; }
      case "prg": { const s = such(M.spalten.prognose, t[3]); return s ? s.ueberschrift : t[3]; }
      case "vk": {
        const s = such(M.spalten.verkauf, t[2]) || such(M.felder.verkaeufe, t[2]);
        return s ? s.ueberschrift : t[2];
      }
      case "rl": { const s = such(M.spalten.ruecklage, t[2]); return s ? s.ueberschrift : t[2]; }
      case "rls": { const s = such(M.spalten.spiegel, t[1]); return s ? s.ueberschrift : t[1]; }
      case "ne": {
        const s = such(M.spalten.neu, t[2]) || such(M.felder.neuobjekte, t[2]);
        return s ? s.ueberschrift : t[2];
      }
      case "neq": return `Quelle ${Number(t[2]) + 1}: ` + ({ rl_geb: "Rücklage Gebäude verfügbar",
        rl_gub: "Rücklage G+B verfügbar", ue1: "ü1 Gebäudegewinn auf Gebäude", ue2: "ü2 G+B-Gewinn auf G+B",
        ue3: "ü3 G+B-Gewinn auf Gebäude", erloes: "Nettoerlös eingesetzt" }[t[3]] || t[3]);
      case "dl": {
        const s = such(M.spalten.darlehen, t[3]);
        const block = { zins: "Zinsen", tilgung: "Tilgung", restschuld: "Restschuld Ende" }[t[3]];
        return (t[2] === "zweit" ? "Zweitkredit: " : "Kredit: ") + (block || (s ? s.ueberschrift : t[3]));
      }
      case "liq": { const s = such(M.spalten.liq, t[2]); return s ? s.ueberschrift : t[2]; }
      case "aus": { const s = such(M.spalten.aus, t[2]); return s ? s.ueberschrift : t[2]; }
      case "vg": return M.vergleich[Number(t[1])].kennzahl;
      case "anl": {
        const s = such(M.spalten.anlage, t[2]) || such(M.felder.anlagen, t[2]);
        return t.length > 3 ? "AfA der Anlage" : s ? s.ueberschrift : t[2];
      }
      default: return pfad;
    }
  }

  function excelName(pfad, M) {
    const t = pfad.split(":");
    const such = (liste, key) => (liste || []).find((x) => x.key === key);
    const name = (s) => (s ? s.name : "");
    switch (t[0]) {
      case "par": return "par_" + t[1];
      case "obj": return name(such(M.felder.objekte, t[2]));
      case "prg": return name(such(M.spalten.prognose, t[3]));
      case "vk": return name(such(M.spalten.verkauf, t[2]) || such(M.felder.verkaeufe, t[2]));
      case "rl": return name(such(M.spalten.ruecklage, t[2]));
      case "rls": return name(such(M.spalten.spiegel, t[1]));
      case "ne": return name(such(M.spalten.neu, t[2]) || such(M.felder.neuobjekte, t[2]));
      case "dl": return { zins: "dl_Zins", tilgung: "dl_Tilgung", restschuld: "dl_Restschuld" }[t[3]]
        || name(such(M.spalten.darlehen, t[3]));
      case "liq": { const s = such(M.spalten.liq, t[2]); return s ? s.name.replace(/^liq_/, (M.szenarien.find((z) => z.key === t[1]) || {}).liq + "_") : ""; }
      case "aus": { const s = such(M.spalten.aus, t[2]); return s ? s.name.replace(/^aus_/, (M.szenarien.find((z) => z.key === t[1]) || {}).aus + "_") : ""; }
      case "vg": return "vg_" + t[2];
      default: return "";
    }
  }

  // ---- Definitionen: je Art und Schlüssel Formel und Teile ----
  // c: {R, t: Pfadteile, i: Jahresindex, jahr, vor(key): Pfad des Vorjahrs, ...}
  // teile: Liste von [Beschriftung, Pfad] oder [Beschriftung, Wert, Format]

  function prgBestand(key, c) {
    const { j, i } = c;
    const p = (k, ii = i) => `prg:b:${j}:${k}:${ii}`;
    const o = (k) => `obj:${j}:${k}`;
    const vorjahr = (k, start) => (i === 0 ? start : p(k, i - 1));
    const steig = (feld, par) => (leer(c.R.objekte[j].felder[feld]) ? `par:${par}` : o(feld));
    const D = {
      aktiv: ["1 bis einschließlich Verkaufsjahr (Verkauf zum Jahresende), danach 0", [["Verkaufsjahr", c.vkJahr]]],
      bestand: ["1 bis zum Jahr vor dem Verkauf, im Verkaufsjahr schon 0", [["Verkaufsjahr", c.vkJahr]]],
      miete: ["Miete bei Halten × aktiv", [["Miete bei Halten", p("miete_halten")], ["aktiv", p("aktiv")]]],
      miete_halten: ["Miete Basisjahr × (1 + Mietsteigerung) ^ Jahre seit dem Basisjahr; ab „Miete eingefroren ab“ keine Steigerung mehr",
        [["Miete Basisjahr", o("miete")], ["Mietsteigerung", steig("mietsteig", "Mietsteig")],
          ["Miete eingefroren ab", o("miete_fix_ab")], ["Basisjahr", "par:Basisjahr"]]],
      einnahmen: ["weitere Einnahmen bei Halten × aktiv", [["bei Halten", p("einnahmen_halten")], ["aktiv", p("aktiv")]]],
      einnahmen_halten: ["weitere Einnahmen Basisjahr × (1 + Mietsteigerung) ^ Jahre seit dem Basisjahr",
        [["weitere Einnahmen Basisjahr", o("weitere_einnahmen")], ["Mietsteigerung", steig("mietsteig", "Mietsteig")]]],
      erhaltung: ["Erhaltung bei Halten × aktiv", [["Erhaltung bei Halten", p("erhaltung_halten")], ["aktiv", p("aktiv")]]],
      erhaltung_halten: ["(Erhaltung Basisjahr × Alterungsfaktor + Großmaßnahme in ihrem Jahr) × (1 + Erhaltungssteigerung) ^ Jahre seit dem Basisjahr. Alterungsfaktor = (1 + zusätzliche Steigerung) ^ Jahre über dem Alter",
        [["Erhaltung Basisjahr", o("erhaltung")], ["Erhaltungssteigerung", steig("erhaltsteig", "Erhaltsteig")],
          ["Baujahr", o("baujahr")], ["Alterung ab Gebäudealter", "par:ErhAlterungAb"],
          ["zusätzliche Steigerung je Jahr", "par:ErhAlterung"], ["Großmaßnahme Jahr", o("san_jahr")],
          ["Großmaßnahme Betrag", o("san_betrag")]]],
      ausgaben: ["weitere Ausgaben Basisjahr × (1 + Kostensteigerung) ^ Jahre seit dem Basisjahr × aktiv",
        [["weitere Ausgaben Basisjahr", o("weitere_ausgaben")], ["Steigerung weitere Ausgaben", "par:Kostensteig"], ["aktiv", p("aktiv")]]],
      afa_halten: ["AfA aus dem AfA-Plan, sonst Summe der AfA je Anlage (Blatt Anlagen), sonst „AfA je Jahr“; höchstens der Buchwert des Vorjahrs bei Halten",
        [["AfA je Jahr (Prognose)", o("afa_jahr")], ["Buchwert Vorjahr bei Halten", vorjahr("buchwert_halten", o("restbuchwert"))],
          ["Anlagen abnutzbar", c.R.objekte[j].anl.anl_abn, "zahl"]]],
      afa: ["AfA bei Halten × aktiv (nach dem Verkaufsjahr 0)", [["AfA bei Halten", p("afa_halten")], ["aktiv", p("aktiv")]]],
      buchwert: ["MAX(Buchwert Vorjahr − AfA; 0)", [["Buchwert Vorjahr", vorjahr("buchwert", o("restbuchwert"))], ["AfA", p("afa")]]],
      buchwert_halten: ["MAX(Buchwert Vorjahr bei Halten − AfA bei Halten; 0)",
        [["Buchwert Vorjahr", vorjahr("buchwert_halten", o("restbuchwert"))], ["AfA bei Halten", p("afa_halten")]]],
      buchwert_gub: ["AK G+B (Grund und Boden wird nicht abgeschrieben)", [["AK G+B", o("ak_gub")]]],
      ergebnis: ["(Miete + weitere Einnahmen − Erhaltung − weitere Ausgaben − AfA) × aktiv",
        [["Miete", p("miete")], ["weitere Einnahmen", p("einnahmen")], ["Erhaltung", p("erhaltung")],
          ["weitere Ausgaben", p("ausgaben")], ["AfA", p("afa")]]],
      verkehrswert: ["Verkehrswert aktuell × (1 + Wertsteigerung) ^ Jahre seit dem Basisjahr",
        [["Verkehrswert aktuell", o("verkehrswert")], ["Wertsteigerung", "par:Wertsteig"]]],
      kredit_zins: N(c.R.objekte[j].felder.restschuld) > 0
        ? ["Restschuld Vorjahresende × Zinssatz", [["Restschuld Vorjahr", vorjahr("restschuld", o("restschuld"))], ["Zinssatz", o("zinssatz")]]]
        : ["ohne Restschuld grob: Zinsaufwand Basisjahr × (1 + Veränderung) ^ Jahre × aktiv",
          [["Zinsaufwand Basisjahr", o("zinsen")], ["Veränderung p. a.", "par:ZinsVeraenderung"], ["aktiv", p("aktiv")]]],
      tilgung: ["MIN(Rate − Zins; Restschuld); im Verkaufsjahr die ganze Restschuld (Ablösung)",
        [["Rate", o("rate")], ["Zinsen", p("kredit_zins")], ["Restschuld Vorjahr", vorjahr("restschuld", o("restschuld"))]]],
      abloesung: ["Tilgung über die planmäßige hinaus im Verkaufsjahr", [["Tilgung", p("tilgung")]]],
      restschuld: ["Restschuld Vorjahr − Tilgung", [["Restschuld Vorjahr", vorjahr("restschuld", o("restschuld"))], ["Tilgung", p("tilgung")]]],
      zins_halten: N(c.R.objekte[j].felder.restschuld) > 0
        ? ["bei Halten: Restschuld Vorjahresende × Zinssatz", [["Restschuld Vorjahr", vorjahr("restschuld_halten", o("restschuld"))], ["Zinssatz", o("zinssatz")]]]
        : ["bei Halten grob: Zinsaufwand Basisjahr × (1 + Veränderung) ^ Jahre", [["Zinsaufwand Basisjahr", o("zinsen")], ["Veränderung p. a.", "par:ZinsVeraenderung"]]],
      tilgung_halten: ["bei Halten: MIN(Rate − Zins; Restschuld)", [["Rate", o("rate")], ["Zinsen", p("zins_halten")], ["Restschuld Vorjahr", vorjahr("restschuld_halten", o("restschuld"))]]],
      restschuld_halten: ["bei Halten: Restschuld Vorjahr − Tilgung", [["Restschuld Vorjahr", vorjahr("restschuld_halten", o("restschuld"))], ["Tilgung", p("tilgung_halten")]]],
      ergebnis_halten: ["bei Halten: Miete + weitere Einnahmen − Erhaltung − weitere Ausgaben − AfA",
        [["Miete", p("miete_halten")], ["weitere Einnahmen", p("einnahmen_halten")], ["Erhaltung", p("erhaltung_halten")],
          ["weitere Ausgaben", o("weitere_ausgaben")], ["AfA", p("afa_halten")]]],
      wirkung: ["Ergebnis nach Darlehenszinsen im Plan − dasselbe bei Halten (Ergebnisbrücke)",
        [["Ergebnis Plan", p("ergebnis")], ["Ergebnis bei Halten", p("ergebnis_halten")], ["Zinsen Plan", p("kredit_zins")], ["Zinsen bei Halten", p("zins_halten")]]],
    };
    return D[key];
  }

  function prgNeu(key, c) {
    const { r, i } = c;
    const p = (k, ii = i) => `prg:n:${r}:${k}:${ii}`;
    const n = (k) => `ne:${r}:${k}`;
    const x = c.R.neuobjekte[r];
    const kst = !leer(x.felder.kst);
    const qm = typeof x.berechnet.miete_qm_jahr === "number";
    const D = {
      aktiv: ["1 ab dem Jahr nach dem Kauf (Kauf zum Jahresende), wenn das Neuobjekt im Modell ist",
        [["Kaufjahr", n("kaufjahr")], ["im Modell", n("gueltig")]]],
      miete: [kst ? "Wert der Kostenstelle Neukauf, sonst wie ohne Kostenstelle"
        : qm ? "Fläche × Mietpreis je m² × 12 × (1 + Mietsteigerung) ^ Jahre seit dem Kauf (ohne die Jahre ohne Steigerung)"
          : "Kaufpreis × Mietrendite × (1 + Mietsteigerung) ^ Jahre seit dem Kauf (ohne die Jahre ohne Steigerung)",
      [["Kaufpreis", n("kaufpreis")], ["Mietrendite", n("mietrendite")], ["Jahresmiete aus m²", n("miete_qm_jahr")],
        ["Mietsteigerung", n("mietsteig_v")], ["Mietsteigerung erst nach Jahren", n("miet_warte")],
        ["Mietsteigerung endet nach Jahren", n("miet_bis")], ["Kaufjahr", n("kaufjahr")], ["aktiv", p("aktiv")]]],
      erhaltung: [qm ? "Miete × Instandhaltung in % der Miete"
        : "Kaufpreis × Erhaltungsquote × Anlauffaktor (in den Anlaufjahren) × (1 + Erhaltungssteigerung) ^ Jahre seit dem Kauf",
      [["Kaufpreis", n("kaufpreis")], ["Erhaltungsquote", n("erhaltungsquote")], ["Anlaufjahre", "par:NeuErhAnlaufJahre"],
        ["Anlauffaktor", "par:NeuErhAnlaufFaktor"], ["Erhaltungssteigerung", "par:Erhaltsteig"], ["Miete", p("miete")],
        ["Instandhaltung in % der Miete", "par:QmErhQuote"]]],
      ausgaben: [qm ? "Miete × Verwaltung in % der Miete" : "Wert der Kostenstelle Neukauf, sonst 0",
        [["Miete", p("miete")], ["Verwaltung in % der Miete", "par:QmVerwQuote"]]],
      afa_qng: ["Bemessungsgrundlage § 7b × Satz, im Kaufjahr und den Folgejahren (höchstens der Buchwert)",
        [["Bemessungsgrundlage § 7b", n("qng_basis")], ["Satz", "par:QNGSatz"], ["Jahre", "par:QNGJahre"]]],
      afa: [String(x.felder.afa_methode).toLowerCase() === "degressiv"
        ? "Sonder-AfA § 7b + Buchwert zu Jahresbeginn × Satz degressiv (ab dem Jahr nach dem Kauf)"
        : "Sonder-AfA § 7b + AfA-Basis × AfA-Satz (ab dem Jahr nach dem Kauf; nach der Sonder-AfA vom Restwert)",
      [["davon Sonder-AfA § 7b", p("afa_qng")], ["AfA-Basis", n("afa_basis")], ["AfA-Satz", n("afa_satz")],
        ["AfA degressiv", "par:AfADegressiv"], ["Buchwert Vorjahr", i ? p("buchwert", i - 1) : ["", 0, "euro"]], ["aktiv", p("aktiv")]]],
      buchwert: ["Kaufjahr: AfA-Basis − AfA; danach MAX(Buchwert Vorjahr − AfA; 0)",
        [["AfA-Basis", n("afa_basis")], ["Buchwert Vorjahr", i ? p("buchwert", i - 1) : ["", 0, "euro"]], ["AfA", p("afa")]]],
      buchwert_gub: ["steuerliche AK G+B (AK G+B − ü2)", [["steuerliche AK G+B", n("ak_gub")]]],
      ergebnis: ["(Miete + weitere Einnahmen − Erhaltung − weitere Ausgaben) × aktiv − AfA",
        [["Miete", p("miete")], ["weitere Einnahmen", p("einnahmen")], ["Erhaltung", p("erhaltung")],
          ["weitere Ausgaben", p("ausgaben")], ["AfA", p("afa")]]],
      verkehrswert: ["Kaufpreis × (1 + Wertsteigerung) ^ Jahre seit dem Kauf", [["Kaufpreis", n("kaufpreis")], ["Wertsteigerung", "par:Wertsteig"]]],
      afa_ohne6b: ["wie die AfA, aber von den vollen AK Gebäude (Szenarien B und C)", [["AK Gebäude", n("ak_geb_neu")], ["AfA-Satz", n("afa_satz")]]],
      afa_qng_ohne6b: ["Sonder-AfA § 7b von der Bemessungsgrundlage ohne § 6b", [["Bemessungsgrundlage ohne § 6b", n("qng_basis_ohne6b")], ["Satz", "par:QNGSatz"]]],
      buchwert_ohne6b: ["wie der Buchwert, aber ab den vollen AK Gebäude", [["AK Gebäude", n("ak_geb_neu")], ["AfA ohne § 6b", p("afa_ohne6b")]]],
      buchwert_gub_ohne6b: ["volle AK G+B ab dem Kaufjahr", [["AK G+B neu", n("ak_gub_neu")]]],
      einnahmen: ["Wert der Kostenstelle Neukauf (weitere Einnahmen), sonst 0", [["Kostenstelle Neukauf", n("kst")]]],
      bestand: ["1 ab dem Kaufjahr, wenn das Neuobjekt im Modell ist", [["Kaufjahr", n("kaufjahr")], ["im Modell", n("gueltig")]]],
      kredit_zins: ["Zinsen von Kredit und Zweitkredit (Blatt Darlehen)", [["Kredit", `dl:${r}:haupt:zins:${i}`], ["Zweitkredit", `dl:${r}:zweit:zins:${i}`]]],
      tilgung: ["Tilgung von Kredit und Zweitkredit (Blatt Darlehen)", [["Kredit", `dl:${r}:haupt:tilgung:${i}`], ["Zweitkredit", `dl:${r}:zweit:tilgung:${i}`]]],
      restschuld: ["Restschuld von Kredit und Zweitkredit (Blatt Darlehen)", [["Kredit", `dl:${r}:haupt:restschuld:${i}`], ["Zweitkredit", `dl:${r}:zweit:restschuld:${i}`]]],
    };
    return D[key];
  }

  function verkauf(key, c) {
    const { j, R } = c;
    const v = (k) => `vk:${j}:${k}`;
    const objNr = R.objekte.findIndex((o) => o && String(o.felder.objekt_id).toLowerCase() === String(R.verkaeufe[j].felder.objekt_id).toLowerCase());
    const i = R.jahre.indexOf(Number(R.verkaeufe[j].felder.jahr));
    const D = {
      preis: R.verkaeufe[j].herkunft.preis === "annahme"
        ? ["Annahme: Verkehrswert × (1 + Wertsteigerung) ^ (Verkaufsjahr − Basisjahr)",
          [["Verkehrswert", `obj:${objNr}:verkehrswert`], ["Wertsteigerung", "par:Wertsteig"], ["Verkaufsjahr", v("jahr")]]]
        : null,
      nettoerloes: ["Verkaufspreis − Verkaufskosten", [["Verkaufspreis", v("preis")], ["Verkaufskosten", v("kosten")]]],
      buchwert_geb: ["Buchwert Gebäude am Ende des Verkaufsjahrs (Prognose, nach dessen AfA)",
        objNr >= 0 && i >= 0 ? [["Buchwert Ende Verkaufsjahr", `prg:b:${objNr}:buchwert:${i}`]] : []],
      ak_gub: ["AK G+B des Objekts", [["AK G+B", `obj:${objNr}:ak_gub`]]],
      quote_gub: ["Preis G+B lt. Kaufvertrag / Preis, sonst Anteil lt. Kaufvertrag, sonst 1 − Verkehrswertanteil Gebäude",
        [["Preis G+B lt. Kaufvertrag", v("gub_betrag")], ["Anteil G+B lt. Kaufvertrag", v("anteil_gub")],
          ["Verkehrswertanteil Gebäude", `obj:${objNr}:vk_quote_gebaeude`]]],
      erloes_gub: ["Nettoerlös × Anteil G+B", [["Nettoerlös", v("nettoerloes")], ["Anteil G+B", v("quote_gub")]]],
      erloes_geb: ["Nettoerlös − Erlösanteil G+B", [["Nettoerlös", v("nettoerloes")], ["Erlösanteil G+B", v("erloes_gub")]]],
      gewinn_geb: ["Erlösanteil Gebäude − Buchwert Gebäude", [["Erlösanteil Gebäude", v("erloes_geb")], ["Buchwert Gebäude", v("buchwert_geb")]]],
      gewinn_gub: ["Erlösanteil G+B − AK G+B", [["Erlösanteil G+B", v("erloes_gub")], ["AK G+B", v("ak_gub")]]],
      gewinn: ["Gewinn Gebäude + Gewinn G+B", [["Gewinn Gebäude", v("gewinn_geb")], ["Gewinn G+B", v("gewinn_gub")]]],
      vorbesitz: ["Verkaufsjahr − Kaufjahr", [["Verkaufsjahr", v("jahr")], ["Kaufjahr", `obj:${objNr}:kaufjahr`], ["Mindestdauer § 6b", "par:6bVorbesitz"]]],
    };
    return D[key];
  }

  function ruecklage(key, c) {
    const { j, R } = c;
    const r = (k) => `rl:${j}:${k}`;
    const nennen = [];
    R.neuobjekte.forEach((x, n) => {
      if (!x.berechnet) return;
      ["quelle", "quelle2", "quelle3"].forEach((g, k) => {
        if (String(x.felder[g]).toLowerCase() === String(R.ruecklagen[j].id).toLowerCase()) nennen.push([n, k]);
      });
    });
    const D = {
      jahr: ["Verkaufsjahr", [["Verkaufsjahr", `vk:${j}:jahr`]]],
      geb: ["MAX(Gewinn Gebäude; 0)", [["Gewinn Gebäude", `vk:${j}:gewinn_geb`]]],
      gub: ["MAX(Gewinn G+B; 0)", [["Gewinn G+B", `vk:${j}:gewinn_gub`]]],
      betrag: ["Rücklage Gebäude + Rücklage G+B", [["Gebäude", r("geb")], ["G+B", r("gub")]]],
      fristjahr: ["Bildungsjahr + Frist (Neubau begonnen: längere Frist)", [["Bildungsjahr", r("jahr")], ["Frist", "par:6bFrist"], ["Frist Neubau", "par:6bFristNeubau"]]],
      ueb_geb: ["Summe ü1 der Neuobjekte, die diese Rücklage nennen", nennen.map(([n, k]) => [R.neuobjekte[n].felder.neu_id + " ü1", `neq:${n}:${k}:ue1`])],
      ueb_gub: ["Summe ü2 + ü3 der Neuobjekte, die diese Rücklage nennen", nennen.flatMap(([n, k]) =>
        [[R.neuobjekte[n].felder.neu_id + " ü2", `neq:${n}:${k}:ue2`], [R.neuobjekte[n].felder.neu_id + " ü3", `neq:${n}:${k}:ue3`]])],
      aufloesung: ["Rücklage − übertragen (Auflösung im Fristjahr)", [["Rücklage", r("betrag")], ["übertragen Gebäude", r("ueb_geb")], ["übertragen G+B", r("ueb_gub")]]],
      zuschlag: ["Auflösung × Gewinnzuschlag × (Fristjahr − Bildungsjahr)", [["Auflösung", r("aufloesung")], ["Zuschlag je Jahr", "par:6bZuschlag"], ["Fristjahr", r("fristjahr")], ["Bildungsjahr", r("jahr")]]],
    };
    return D[key];
  }

  function spiegel(key, c) {
    const { i, R } = c;
    const s = (k, ii = i) => `rls:${k}:${ii}`;
    const jahr = R.jahre[i];
    const D = {
      gewinne: ["Summe der Veräußerungsgewinne im Jahr", R.verkaeufe.map((v, j) => [v, j])
        .filter(([v]) => v.gueltig && Number(v.felder.jahr) === jahr).map(([v, j]) => [v.felder.objekt_id, `vk:${j}:gewinn`])],
      bildung: ["Summe der Rücklagen mit Bildungsjahr im Jahr", R.ruecklagen.map((r, j) => [r, j])
        .filter(([r]) => r && Number(r.jahr) === jahr).map(([r, j]) => [r.id, `rl:${j}:betrag`])],
      aufloesung: ["Summe der Auflösungen mit Fristjahr im Jahr", R.ruecklagen.map((r, j) => [r, j])
        .filter(([r]) => r && r.fristjahr === jahr).map(([r, j]) => [r.id, `rl:${j}:aufloesung`])],
      zuschlag: ["Summe der Gewinnzuschläge mit Fristjahr im Jahr", R.ruecklagen.map((r, j) => [r, j])
        .filter(([r]) => r && r.fristjahr === jahr).map(([r, j]) => [r.id, `rl:${j}:zuschlag`])],
      uebertragung: ["Übertragungen der Neuobjekte mit Kaufjahr im Jahr (Kauf im Vorjahr der Veräußerung: im Verkaufsjahr)",
        R.neuobjekte.map((x, n) => [x, n]).filter(([x]) => x.berechnet && Number(x.felder.kaufjahr) === jahr)
          .map(([x, n]) => [x.felder.neu_id, `ne:${n}:ue_gesamt`])],
      steuerwirksam: ["Gewinne − Einstellung + Auflösung + Zuschlag",
        [["Gewinne", s("gewinne")], ["Einstellung", s("bildung")], ["Auflösung", s("aufloesung")], ["Zuschlag", s("zuschlag")]]],
      bestand: ["Bestand Vorjahr + Einstellung − Übertragung − Auflösung",
        [["Bestand Vorjahr", i ? s("bestand", i - 1) : ["", 0, "euro"]], ["Einstellung", s("bildung")],
          ["Übertragung", s("uebertragung")], ["Auflösung", s("aufloesung")]]],
    };
    return D[key];
  }

  function neu(key, c) {
    const { r, R } = c;
    const n = (k) => `ne:${r}:${k}`;
    const q = (k) => [0, 1, 2].map((kk) => [`Quelle ${kk + 1}`, `neq:${r}:${kk}:${k}`]);
    const D = {
      ak_gub_neu: ["AK G+B eingetragen, sonst (Kaufpreis + Nebenkosten) × Anteil G+B",
        [["Kaufpreis", n("kaufpreis")], ["Nebenkosten", n("nebenkosten")], ["Anteil G+B", n("anteil_gub")], ["AK G+B (Eingabe)", n("ak_gub_eingabe")]]],
      ak_geb_neu: ["AK Gebäude eingetragen, sonst Kaufpreis + Nebenkosten − AK G+B",
        [["Kaufpreis", n("kaufpreis")], ["Nebenkosten", n("nebenkosten")], ["AK G+B neu", n("ak_gub_neu")]]],
      ue1: ["Summe ü1 über die Quellen", q("ue1")],
      ue2: ["Summe ü2 über die Quellen", q("ue2")],
      ue3: ["Summe ü3 über die Quellen", q("ue3")],
      ue_gesamt: ["ü1 + ü2 + ü3", [["ü1", n("ue1")], ["ü2", n("ue2")], ["ü3", n("ue3")]]],
      afa_basis: ["AK Gebäude neu − ü1 − ü3", [["AK Gebäude neu", n("ak_geb_neu")], ["ü1", n("ue1")], ["ü3", n("ue3")]]],
      ak_gub: ["AK G+B neu − ü2", [["AK G+B neu", n("ak_gub_neu")], ["ü2", n("ue2")]]],
      qng_basis: ["AfA-Basis, mit Fläche höchstens 4.000 € je m²; 0 ohne QNG oder über der Baukostenobergrenze",
        [["AfA-Basis", n("afa_basis")], ["Fläche m²", n("flaeche")], ["höchstens je m²", "par:QNGHoechstQm"], ["Sonder-AfA § 7b", n("qng")]]],
      ak_gesamt: ["Kaufpreis + Nebenkosten", [["Kaufpreis", n("kaufpreis")], ["Nebenkosten", n("nebenkosten")]]],
      erloes: ["eingesetzter Nettoerlös der Quell-Verkäufe (nur mit „Verkaufserlös zur Finanzierung = ja“)", q("erloes")],
      bedarf: ["Kaufpreis + Nebenkosten − eingesetzter Verkaufserlös", [["AK gesamt", n("ak_gesamt")], ["Verkaufserlös", n("erloes")]]],
      kredit: [!leer(R.neuobjekte[r].felder.kredit_betrag) ? "Kreditbetrag (Eingabe)"
        : !leer(R.neuobjekte[r].felder.kredit_quote) ? "(Kaufpreis + Nebenkosten) × Kreditquote" : "ganzer Finanzierungsbedarf",
      [["Finanzierung Rest", n("fin_art")], ["Kreditbetrag", n("kredit_betrag")], ["Kreditquote", n("kredit_quote")],
        ["AK gesamt", n("ak_gesamt")], ["Finanzierungsbedarf", n("bedarf")]]],
      eigen: ["Finanzierungsbedarf − Kredit", [["Finanzierungsbedarf", n("bedarf")], ["Kredit", n("kredit")]]],
    };
    return D[key];
  }

  function neuQuelle(key, c) {
    const { r, k, R } = c;
    const x = R.neuobjekte[r];
    const qid = x.felder[["quelle", "quelle2", "quelle3"][k]];
    const rlNr = R.ruecklagen.findIndex((z) => z && String(z.id).toLowerCase() === String(qid).toLowerCase());
    const z = (kk) => `neq:${r}:${k}:${kk}`;
    const D = {
      rl_geb: ["Rücklage Gebäude der Quelle − was Neuobjekte in den Zeilen darüber schon als ü1 genommen haben",
        rlNr >= 0 ? [["Rücklage Gebäude", `rl:${rlNr}:geb`]] : []],
      rl_gub: ["Rücklage G+B der Quelle − was Zeilen darüber schon als ü2 und ü3 genommen haben",
        rlNr >= 0 ? [["Rücklage G+B", `rl:${rlNr}:gub`]] : []],
      ue1: ["MIN(verfügbare Rücklage Gebäude; AK Gebäude bzw. QNG-Grenze − ü1 der Quellen davor)",
        [["Rücklage Gebäude verfügbar", z("rl_geb")], ["AK Gebäude neu", `ne:${r}:ak_geb_neu`], ["§ 6b nur bis QNG-Grenze", `ne:${r}:qng_grenze`]]],
      ue2: ["MIN(verfügbare Rücklage G+B; AK G+B − ü2 der Quellen davor)", [["Rücklage G+B verfügbar", z("rl_gub")], ["AK G+B neu", `ne:${r}:ak_gub_neu`]]],
      ue3: ["MIN(Rücklage G+B − ü2; Rest der AK Gebäude nach ü1)", [["Rücklage G+B verfügbar", z("rl_gub")], ["ü2", z("ue2")], ["ü1 gesamt", `ne:${r}:ue1`]]],
      erloes: ["MIN(Nettoerlös der Quelle − schon eingesetzt; AK gesamt − fester Kredit)",
        rlNr >= 0 ? [["Nettoerlös", `vk:${rlNr}:nettoerloes`], ["AK gesamt", `ne:${r}:ak_gesamt`]] : []],
    };
    return D[key];
  }

  function darlehen(key, c) {
    const { r, teil, i } = c;
    const d = (k, ii) => (ii === undefined ? `dl:${r}:${teil}:${k}` : `dl:${r}:${teil}:${k}:${ii}`);
    const vor = i > 0 ? d("restschuld", i - 1) : ["Restschuld Vorjahr", 0, "euro"];
    const D = {
      zins: ["Restschuld Vorjahresende × Zinssatz", [["Restschuld Vorjahr", vor], ["Zinssatz", d("zinssatz")]]],
      tilgung: ["Annuität: Rate − Zins; linear: Tilgung p. a. × Kreditbetrag; endfällig: nach der Laufzeit alles. In tilgungsfreien Jahren 0, höchstens die Restschuld",
        [["Rate", d("rate")], ["Zinsen", d("zins", i)], ["Tilgungsart", d("art")], ["Laufzeit", d("laufzeit")],
          ["tilgungsfreie Jahre", d("tilgungsfrei")], ["Restschuld Vorjahr", vor]]],
      restschuld: ["Kaufjahr: Kreditbetrag; danach Restschuld Vorjahr − Tilgung", [["Kreditbetrag", d("betrag")], ["Restschuld Vorjahr", vor], ["Tilgung", d("tilgung", i)]]],
      rate: ["Annuität: Kreditbetrag × (Zinssatz + Tilgung); linear: Kreditbetrag × Tilgung", [["Kreditbetrag", d("betrag")], ["Zinssatz", d("zinssatz")], ["Tilgung p. a.", d("tilgungssatz")]]],
      betrag: [teil === "zweit" ? "Zweitkredit: Betrag, sonst Kredit × Anteil" : "Kredit − Zweitkredit",
        [["Kredit gesamt", `ne:${r}:kredit`], ["davon Zweitkredit", `ne:${r}:kredit2`]]],
    };
    return D[key];
  }

  function liq(key, c) {
    const { sz, i, R } = c;
    const l = (k, ii = i) => `liq:${sz}:${k}:${ii}`;
    const summeObjekte = (k, kNeu = k) => {
      const teile = [];
      R.prognose.bestand.forEach((s, j) => { if (s && (sz !== "B" || s.mit_quelle[i] === 0)) teile.push([s.id[0], `prg:b:${j}:${k}:${i}`]); });
      if (sz !== "Baseline") R.prognose.neu.forEach((s, r) => {
        if (s && (sz !== "B" || s.mit_quelle[i] === 0)) teile.push([s.id[0], `prg:n:${r}:${kNeu}:${i}`]);
      });
      return teile;
    };
    const halten = sz === "Baseline";
    const D = {
      einnahmen: [halten ? "Summe Miete und weitere Einnahmen bei Halten" : "Summe Miete und weitere Einnahmen aller Objekte",
        [...summeObjekte(halten ? "miete_halten" : "miete"), ...summeObjekte(halten ? "einnahmen_halten" : "einnahmen")]],
      ausgaben: ["Summe Erhaltung und weitere Ausgaben aller Objekte",
        [...summeObjekte(halten ? "erhaltung_halten" : "erhaltung"), ...summeObjekte("ausgaben")]],
      afa: [sz === "A" ? "Summe AfA aller Objekte" : halten ? "Rückgang des Buchwerts bei Halten" : "Summe AfA ohne § 6b-Übertragung",
        summeObjekte(sz === "A" ? "afa" : halten ? "afa_halten" : "afa_ohne6b")],
      ergebnis: ["Einnahmen − Ausgaben − AfA", [["Einnahmen", l("einnahmen")], ["Ausgaben", l("ausgaben")], ["AfA", l("afa")]]],
      verkauf: [sz === "A" ? "steuerwirksam aus Verkauf und Rücklage (Rücklagenspiegel)" : halten ? "kein Verkauf" : "jeder Veräußerungsgewinn sofort",
        halten ? [] : [[sz === "A" ? "steuerwirksam" : "Gewinne", `rls:${sz === "A" ? "steuerwirksam" : "gewinne"}:${i}`]]],
      zins: ["Liquidität Vorjahresende × Rendite Alternativanlage",
        [["Liquidität Vorjahr", i ? l("kum", i - 1) : ["", 0, "euro"]], ["Rendite", "par:Alternativrendite"]]],
      kreditzins: ["Summe Zinsen der Darlehen", summeObjekte(halten ? "zins_halten" : "kredit_zins")],
      zve: ["laufendes Ergebnis + Verkauf + Zinsertrag − Zinsen Darlehen",
        [["laufendes Ergebnis", l("ergebnis")], ["aus Verkauf", l("verkauf")], ["Zinsertrag", l("zins")], ["Zinsen Darlehen", l("kreditzins")]]],
      vortrag_genutzt: ["MIN(Verlustvortrag Vorjahr; positives Ergebnis)", [["Vortrag Vorjahr", i ? l("vortrag", i - 1) : ["", 0, "euro"]], ["Ergebnis vor Vortrag", l("zve")]]],
      bemessung: ["positives Ergebnis − genutzter Verlustvortrag", [["Ergebnis vor Vortrag", l("zve")], ["Vortrag genutzt", l("vortrag_genutzt")]]],
      vortrag: ["Vortrag Vorjahr − genutzt + Verlust des Jahres", [["Vortrag Vorjahr", i ? l("vortrag", i - 1) : ["", 0, "euro"]], ["genutzt", l("vortrag_genutzt")], ["Ergebnis vor Vortrag", l("zve")]]],
      steuer: ["Bemessungsgrundlage × angewandter Steuersatz (+ Verlust × Satz auf Verluste)",
        [["Bemessungsgrundlage", l("bemessung")], ["angewandter Steuersatz", "par:SteuersatzEff"], ["Satz auf Verluste", "par:VerlustSatz"]]],
      verkaufserloes: ["Summe Nettoerlöse der Verkäufe im Jahr", halten ? [] : R.verkaeufe.map((v, j) => [v, j])
        .filter(([v]) => v.gueltig && Number(v.felder.jahr) === R.jahre[i]).map(([v, j]) => [v.felder.objekt_id, `vk:${j}:nettoerloes`])],
      kauf: ["Kaufpreis + Nebenkosten der Neuobjekte im Kaufjahr", halten ? [] : R.neuobjekte.map((x, r) => [x, r])
        .filter(([x]) => x.berechnet && x.berechnet.gueltig === 1 && Number(x.felder.kaufjahr) === R.jahre[i] && (sz !== "B" || x.berechnet.mit_quelle === 0))
        .map(([x, r]) => [x.felder.neu_id, `ne:${r}:ak_gesamt`])],
      kredit: ["Kreditauszahlung zum Kauf", halten ? [] : R.neuobjekte.map((x, r) => [x, r])
        .filter(([x]) => x.berechnet && x.berechnet.gueltig === 1 && Number(x.felder.kaufjahr) === R.jahre[i] && (sz !== "B" || x.berechnet.mit_quelle === 0))
        .map(([x, r]) => [x.felder.neu_id, `ne:${r}:kredit`])],
      tilgung: ["Summe Tilgung (mit Ablösung beim Verkauf)", summeObjekte(halten ? "tilgung_halten" : "tilgung")],
      abloesung: ["Summe der Ablösungen von Bestandsdarlehen im Verkaufsjahr", halten ? [] : summeObjekte("abloesung")],
      rueckfluss: ["Verkaufserlöse − Veräußerungsgewinne: der steuerneutrale Teil (Buchwert Gebäude + AK G+B)",
        [["Verkaufserlöse", l("verkaufserloes")], ["Gewinne", `rls:gewinne:${i}`]]],
      ausschuettung: ["Ausschüttung p. a. × (1 + Steigerung) ^ Jahre seit Beginn",
        [["Ausschüttung p. a.", "par:Ausschuettung"], ["ab Jahr", "par:AusschuettungAb"], ["Steigerung", "par:AusschuettungSteig"]]],
      zufluss: ["Einnahmen − Ausgaben + Zinsertrag − Zinsen Darlehen + Verkaufserlöse − Steuer − Kauf + Kredit − Tilgung − Ausschüttung",
        [["Einnahmen", l("einnahmen")], ["Ausgaben", l("ausgaben")], ["Zinsertrag", l("zins")], ["Zinsen Darlehen", l("kreditzins")],
          ["Verkaufserlöse", l("verkaufserloes")], ["Steuer", l("steuer")], ["Kauf", l("kauf")], ["Kredit", l("kredit")],
          ["Tilgung", l("tilgung")], ["Ausschüttung", l("ausschuettung")]]],
      kum: ["Liquidität Vorjahr + freier Mittelzufluss", [["Liquidität Vorjahr", i ? l("kum", i - 1) : ["", 0, "euro"]], ["Mittelzufluss", l("zufluss")]]],
      restschuld: ["Summe Restschuld aller Darlehen", summeObjekte(halten ? "restschuld_halten" : "restschuld")],
    };
    return D[key];
  }

  function aus(key, c) {
    const { sz, i, R } = c;
    const a = (k, ii = i) => `aus:${sz}:${k}:${ii}`;
    const l = (k) => `liq:${sz}:${k}:${i}`;
    const imBestand = [];
    const geb = sz === "Baseline" ? "buchwert_halten" : sz === "A" ? "buchwert" : "buchwert_ohne6b";
    const gub = sz === "A" || sz === "Baseline" ? "buchwert_gub" : "buchwert_gub_ohne6b";
    R.prognose.bestand.forEach((s, j) => { if (s && (sz === "Baseline" || (s.bestand[i] === 1 && (sz !== "B" || s.mit_quelle[i] === 0)))) imBestand.push(["b", j, s]); });
    if (sz !== "Baseline") R.prognose.neu.forEach((s, r) => { if (s && s.bestand[i] === 1 && (sz !== "B" || s.mit_quelle[i] === 0)) imBestand.push(["n", r, s]); });
    const D = {
      guv: ["laufendes Ergebnis + Verkauf + Zinsertrag − Zinsen Darlehen", [["laufendes Ergebnis", a("ergebnis")], ["Verkauf", a("verkauf")], ["Zinsertrag", a("zins")], ["Zinsen Darlehen", a("kreditzins")]]],
      nach_steuer: ["Gesamt-GuV − Steuer", [["Gesamt-GuV", a("guv")], ["Steuer", a("steuer")]]],
      steuer_kum: ["Steuer kumuliert Vorjahr + Steuer", [["Vorjahr", i ? a("steuer_kum", i - 1) : ["", 0, "euro"]], ["Steuer", a("steuer")]]],
      verkehrswert: ["Summe Verkehrswert der Objekte im Bestand am Jahresende", imBestand.map(([t, n, s]) => [s.id[0], `prg:${t}:${n}:verkehrswert:${i}`])],
      buchwert: ["Summe Buchwert Gebäude + G+B der Objekte im Bestand", imBestand.flatMap(([t, n, s]) => [[s.id[0] + " Gebäude", `prg:${t}:${n}:${geb}:${i}`], [s.id[0] + " G+B", `prg:${t}:${n}:${gub}:${i}`]])],
      stille_reserven: ["Verkehrswert − Buchwert", [["Verkehrswert", a("verkehrswert")], ["Buchwert", a("buchwert")]]],
      ruecklage: [sz === "A" ? "Rücklagenbestand Ende (Rücklagenspiegel)" : "keine Rücklage in diesem Szenario", sz === "A" ? [["Bestand", `rls:bestand:${i}`]] : []],
      liquiditaet: ["Liquidität kumuliert (Blatt Liquidität)", [["Liquidität", l("kum")]]],
      restschuld: ["Restschuld der Darlehen (Blatt Liquidität)", [["Restschuld", l("restschuld")]]],
      vermoegen: ["Verkehrswert + Liquidität − Restschuld", [["Verkehrswert", a("verkehrswert")], ["Liquidität", a("liquiditaet")], ["Restschuld", a("restschuld")]]],
      latente_steuer: ["(stille Reserven + Rücklage − Verlustvortrag) × angewandter Steuersatz: Steuer, wenn alles zum Verkehrswert verkauft würde",
        [["stille Reserven", a("stille_reserven")], ["Rücklage", a("ruecklage")], ["Verlustvortrag", a("vortrag")], ["Steuersatz", "par:SteuersatzEff"]]],
      vermoegen_netto: ["Gesamtvermögen vor latenter Steuer − latente Steuer", [["vor latenter Steuer", a("vermoegen")], ["latente Steuer", a("latente_steuer")]]],
    };
    for (const k of ["ergebnis", "verkauf", "zins", "kreditzins", "steuer", "vortrag"]) {
      if (!D[k]) D[k] = ["aus dem Blatt Liquidität", [["Liquidität", l(k === "liquiditaet" ? "kum" : k)]]];
    }
    return D[key];
  }

  function objekt(key, c) {
    const { j, R } = c;
    const o = (k) => `obj:${j}:${k}`;
    const herkunft = R.objekte[j].herkunft[key];
    if (herkunft === "anlagen") return ["aus dem Blatt Anlagen (Summe je Objekt)", []];
    if (herkunft === "kostenstelle") return ["aus dem Kostenstellenblatt (Basisjahr)", []];
    if (herkunft !== "annahme") return null;
    const D = {
      erhaltung: ["Annahme: Miete × Erhaltungsquote", [["Miete", o("miete")], ["Erhaltung in % der Miete", "par:AnnErhQuote"]]],
      verkehrswert: ["Annahme: (Miete + weitere Einnahmen) × Vervielfältiger", [["Miete", o("miete")], ["weitere Einnahmen", o("weitere_einnahmen")], ["Vervielfältiger", "par:AnnVervielfaeltiger"]]],
      vk_quote_gebaeude: ["Annahme: Gebäudeanteil", [["Gebäudeanteil", "par:AnnGebaeudeanteil"]]],
      afa_satz: ["Annahme: AfA-Satz Bestand", [["AfA-Satz Bestand", "par:AnnAfASatz"]]],
      kaufjahr: ["Annahme: Basisjahr − Jahre seit Kauf", [["Basisjahr", "par:Basisjahr"], ["Jahre seit Kauf", "par:AnnHaltedauer"]]],
      ak_gebaeude: ["Annahme: AfA lt. Buchhaltung / AfA-Satz, sonst Verkehrswert × Gebäudeanteil, abgezinst bis zum Kaufjahr",
        [["AfA lt. Buchhaltung", o("afa_bwa")], ["AfA-Satz", o("afa_satz")], ["Verkehrswert", o("verkehrswert")],
          ["Gebäudeanteil", o("vk_quote_gebaeude")], ["Wertsteigerung", "par:Wertsteig"], ["Kaufjahr", o("kaufjahr")]]],
      ak_gub: ["Annahme: AK Gebäude × (1 − Gebäudeanteil) / Gebäudeanteil", [["AK Gebäude", o("ak_gebaeude")], ["Gebäudeanteil", o("vk_quote_gebaeude")]]],
      restbuchwert: ["Annahme: AK Gebäude × (1 − Jahre seit Kauf × AfA-Satz); 0, wenn die Buchhaltung keine AfA mehr zeigt",
        [["AK Gebäude", o("ak_gebaeude")], ["Kaufjahr", o("kaufjahr")], ["AfA-Satz", o("afa_satz")], ["AfA lt. Buchhaltung", o("afa_bwa")]]],
      afa_jahr: ["AfA lt. Buchhaltung, sonst AK Gebäude × AfA-Satz", [["AfA lt. Buchhaltung", o("afa_bwa")], ["AK Gebäude", o("ak_gebaeude")], ["AfA-Satz", o("afa_satz")]]],
      mietsteig: ["Parameter Mietsteigerung", [["Mietsteigerung", "par:Mietsteig"]]],
      erhaltsteig: ["Parameter Erhaltungssteigerung", [["Erhaltungssteigerung", "par:Erhaltsteig"]]],
      san_jahr: ["Annahme: Baujahr + Alter für Großmaßnahme, frühestens erstes Prognosejahr + Vorlauf; ohne Baujahr keine",
        [["Baujahr", o("baujahr")], ["fällig bei Gebäudealter", "par:SanAlter"], ["Vorlauf", "par:SanVorlauf"], ["Quote", "par:SanQuote"]]],
      san_betrag: ["Annahme: Verkehrswert × Gebäudeanteil × Quote", [["Verkehrswert", o("verkehrswert")], ["Gebäudeanteil", o("vk_quote_gebaeude")], ["Quote", "par:SanQuote"]]],
      zinssatz: ["Annahme: Zinsaufwand Basisjahr / Restschuld", [["Zinsaufwand", o("zinsen")], ["Restschuld", o("restschuld")]]],
      rate: ["Annahme: Restschuld × (Zinssatz + Tilgung)", [["Restschuld", o("restschuld")], ["Zinssatz", o("zinssatz")], ["Tilgung", "par:AnnDarlTilgung"]]],
    };
    return D[key] || null;
  }

  function vergleich(c) {
    const { k, sz, R } = c;
    const z = R.vergleich[k];
    const letzte = R.jahre.length - 1;
    if (z.art === "ende") return [`Wert im letzten Prognosejahr ${R.jahre[letzte]}`, [[String(R.jahre[letzte]), `${z.tabelle}:${sz}:${z.key}:${letzte}`]]];
    return ["Summe über alle Prognosejahre", R.jahre.map((t, i) => [String(t), `${z.tabelle}:${sz}:${z.key}:${i}`])];
  }

  /** Rechenweg zu einem Pfad: {titel, wer, jahr, wert, formel, teile, excel}. */
  function erklaere(pfad, R, M) {
    const t = pfad.split(":");
    const z = (x) => Number(x);
    let def = null, jahr = null;
    try {
      switch (t[0]) {
        case "prg": {
          const i = z(t[4]);
          jahr = R.jahre[i];
          if (t[1] === "b") {
            const j = z(t[2]);
            const vk = R.verkaeufe.findIndex((v) => v.aktiv && String(v.umfang).toLowerCase()
              .includes(";" + String(R.objekte[j].felder.objekt_id).toLowerCase() + ";") && v.felder.jahr !== "");
            def = prgBestand(t[3], { R, j, i, vkJahr: vk >= 0 ? `vk:${vk}:jahr` : ["kein Verkauf", "", "text"] });
          } else def = prgNeu(t[3], { R, r: z(t[2]), i });
          break;
        }
        case "vk": def = verkauf(t[2], { R, j: z(t[1]) }); break;
        case "rl": def = ruecklage(t[2], { R, j: z(t[1]) }); break;
        case "rls": jahr = R.jahre[z(t[2])]; def = spiegel(t[1], { R, i: z(t[2]) }); break;
        case "ne": def = neu(t[2], { R, r: z(t[1]) }); break;
        case "neq": def = neuQuelle(t[3], { R, r: z(t[1]), k: z(t[2]) }); break;
        case "dl": if (t.length > 4) jahr = R.jahre[z(t[4])];
          def = darlehen(t[3], { R, r: z(t[1]), teil: t[2], i: t.length > 4 ? z(t[4]) : 0 }); break;
        case "liq": jahr = R.jahre[z(t[3])]; def = liq(t[2], { R, sz: t[1], i: z(t[3]) }); break;
        case "aus": jahr = R.jahre[z(t[3])]; def = aus(t[2], { R, sz: t[1], i: z(t[3]) }); break;
        case "vg": def = vergleich({ R, k: z(t[1]), sz: t[2] }); break;
        case "obj": def = objekt(t[2], { R, j: z(t[1]) }); break;
        default: def = null;
      }
    } catch (e) {
      def = null;
    }
    const teile = def ? def[1].map((teil) => {
      const [label, ref, fmt] = teil;
      if (Array.isArray(ref)) return { label: label || ref[0], wert: ref[1], format: ref[2] };
      if (typeof ref === "string" && /^(par|obj|prg|vk|rl|rls|ne|neq|dl|liq|aus|vg|anl):/.test(ref))
        return { label, pfad: ref, wert: wertVon(ref, R), wer: wer(ref, R) };
      return { label, wert: ref, format: fmt || "euro" };
    }) : [];
    return {
      pfad, titel: bezeichnung(pfad, M), wer: wer(pfad, R), jahr, wert: wertVon(pfad, R),
      formel: def ? def[0] : null, teile, excel: excelName(pfad, M),
    };
  }

  return { erklaere, wertVon, bezeichnung, excelName, wer };
});
