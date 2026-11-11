/*
 * Optimierung und Varianten der Browser-Fassung: das Gegenstück zu den Makros
 * modOptimierung und modVarianten der .xlsm.
 *
 * Wie die Makros rechnet dieses Modul nicht selbst. Es setzt Eingaben, lässt das Rechenwerk
 * rechnen und liest Ziel und Nebenbedingungen aus dem Ergebnis:
 *   Ziel            Endvermögen nach latenter Steuer, Plan A (opt_Ziel)
 *   zulässig, wenn  kein Jahr unter der Mindestliquidität (opt_LiqJahre = 0), kein Jahr unter
 *                   der Mindest-Kapitaldienstdeckung (opt_KDDJahre = 0), keine neuen Fehler
 * Suche: koordinatenweise, je Runde jeder Hebel jedes Neuobjekts einmal; ein Kandidat bleibt
 * nur, wenn er um mehr als 1 € besser ist. Unzulässige Lösungen werden nach Abstand zur
 * Zulässigkeit verglichen, damit die Suche auch aus einem unzulässigen Ausgang herausfindet.
 *
 * Läuft im Browser (window.Optimierung) und in Node (require), damit
 * pruefung/pruefen_optimierung.py dieselbe Datei gegen die Sollwerte der Makrofälle prüft.
 */
(function (wurzel, fabrik) {
  if (typeof module === "object" && module.exports) module.exports = fabrik(require("./rechenwerk.js"));
  else wurzel.Optimierung = fabrik(wurzel.Rechenwerk);
})(typeof self !== "undefined" ? self : this, function (RW) {
  "use strict";

  const UNZULAESSIG = -1e15;
  const BESSER_AB = 1;          // Euro: kleinere Verbesserungen zählen nicht
  const ARTEN = ["jahr", "qng", "quote"];
  const HEBEL_NAME = { jahr: "Kaufjahr", qng: "§ 6b nur bis QNG-Grenze", quote: "Kreditquote" };
  const MAX_VARIANTEN = 100;

  const leer = (v) => v === null || v === undefined || v === "";
  const istZahl = (v) => typeof v === "number" && isFinite(v);
  const N = (v) => (istZahl(v) ? v : v === true ? 1 : 0);
  const klein = (v) => String(leer(v) ? "" : v).trim().toLowerCase();
  const kopie = (x) => JSON.parse(JSON.stringify(x));
  const summe = (liste) => liste.reduce((a, b) => a + N(b), 0);

  /** Einstellungen der Suche mit den Vorgaben des Blatts Optimierung (OPT_EINSTELLUNGEN). */
  const EINSTELLUNGEN = [
    { key: "hebelJahr", name: "opt_HebelJahr", text: "Kaufjahr verändern", wert: "ja", format: "text", auswahl: ["ja", "nein"],
      hinweis: "ja: Kaufjahr je Neuobjekt um bis zu ± Spielraum; mit Quelle nur innerhalb der Frist der Rücklage (sonst Fehler, verworfen)" },
    { key: "jahrSpielraum", name: "opt_JahrSpielraum", text: "Spielraum Kaufjahr ± Jahre", wert: 2, format: "zahl",
      hinweis: "Jahre vor und nach dem eingetragenen Kaufjahr, innerhalb des Rasters" },
    { key: "hebel6b", name: "opt_Hebel6b", text: "§ 6b nur bis QNG-Grenze verändern", wert: "ja", format: "text", auswahl: ["ja", "nein"],
      hinweis: "ja: probiert ja und nein bei Neuobjekten mit Sonder-AfA § 7b (QNG) und Quelle" },
    { key: "hebelKredit", name: "opt_HebelKredit", text: "Kreditquote verändern", wert: "ja", format: "text", auswahl: ["ja", "nein"],
      hinweis: "ja: Finanzierung Rest = Kredit mit Kreditquote von … bis …; braucht Zinssatz und Tilgung (bzw. Laufzeit) beim Neuobjekt, sonst bleibt es wie erfasst" },
    { key: "kreditMin", name: "opt_KreditMin", text: "Kreditquote von", wert: 0, format: "prozent", hinweis: "kleinste geprüfte Quote" },
    { key: "kreditMax", name: "opt_KreditMax", text: "Kreditquote bis", wert: 0.8, format: "prozent",
      hinweis: "größte geprüfte Quote, z. B. Beleihungsgrenze der Bank" },
    { key: "kreditSchritt", name: "opt_KreditSchritt", text: "Schrittweite Kreditquote", wert: 0.2, format: "prozent",
      hinweis: "kleiner = genauer, aber mehr Rechenläufe" },
    { key: "runden", name: "opt_Runden", text: "Runden höchstens", wert: 3, format: "zahl",
      hinweis: "je Runde jeder Hebel jedes Neuobjekts einmal; Ende, sobald eine Runde nichts verbessert" },
    { key: "festhalten", name: "opt_Festhalten", text: "Ausgang und Ergebnis als Variante festhalten", wert: "ja", format: "text", auswahl: ["ja", "nein"],
      hinweis: "ja: Varianten „Optimierung Ausgang“ und „Optimierung Ergebnis“; der Ausgang lässt sich wieder einspielen" },
  ];
  function einstellungen(roh) {
    const e = {};
    for (const s of EINSTELLUNGEN) e[s.key] = roh && !leer(roh[s.key]) ? roh[s.key] : s.wert;
    e.ohne = Array.isArray(roh && roh.ohne) ? roh.ohne.slice() : [];   // NeuIDs mit „optimieren = nein“
    return e;
  }

  // ---- Kennzahlen (Blatt Optimierung, formeln.optimierung_kennzahlen) ----
  /** fehlerKeys: Schlüssel der Prüfungen mit Art „Fehler“ (pr_Fehler zählt die auffälligen). */
  function kennzahlen(R, fehlerKeys) {
    const P = R.parameter, L = R.szenarien.A.liq;
    const mindest = N(P.MindestLiq);
    let minI = 0;
    L.kum.forEach((w, i) => { if (w < L.kum[minI]) minI = i; });
    const mitKDD = istZahl(P.MindestKDD);
    const kddJahre = mitKDD ? L.kum.filter((_, i) => {
      const kd = L.kreditzins[i] + L.tilgung[i] - L.abloesung[i];
      return kd > 0.005 && L.einnahmen[i] - L.ausgaben[i] < P.MindestKDD * kd;
    }).length : 0;
    const liqJahre = L.kum.filter((k) => k < mindest - 0.005).length;
    return {
      ziel: R.vergleich[0].A,
      zielHalten: R.vergleich[0].Baseline,
      ausschuettung: N(P.Ausschuettung),
      ausschuettungSumme: summe(L.ausschuettung),
      mindestLiq: mindest,
      liqMin: L.kum.length ? L.kum[minI] : 0,
      liqMinJahr: R.jahre[minI],
      liqJahre,
      fehlbetrag: summe(L.kum.map((k) => (k < mindest ? mindest - k : 0))),
      mindestKDD: mitKDD ? P.MindestKDD : "keine",
      kddJahre,
      erfuellt: liqJahre === 0 && kddJahre === 0,
      fehler: (fehlerKeys || []).filter((k) => N(R.pruefung[k]) > 0).length,
    };
  }

  /** Text wie pr_Gesamt: „2 Fehler“, „1 Warnung(en)“ oder „OK“. */
  function pruefungText(R, pruefungen) {
    const zaehle = (art) => (pruefungen || []).filter((p) => p.art === art && N(R.pruefung[p.key]) > 0).length;
    const fehler = zaehle("Fehler"), warn = zaehle("Warnung");
    return fehler ? `${fehler} Fehler` : warn ? `${warn} Warnung(en)` : "OK";
  }

  // ---- Neuobjekte: aktuelle Hebel ----
  function quoteVon(x) {
    const b = x.berechnet;
    return b && N(b.ak_gesamt) !== 0 ? N(b.kredit) / b.ak_gesamt : "";
  }
  /** Zeilen der Tabelle „Neuobjekte“ im Blatt Optimierung: aktuelle Hebel je Neuobjekt. */
  function neuobjektZeilen(R, einst) {
    const e = einstellungen(einst);
    return R.neuobjekte.map((x, r) => {
      const f = x.felder;
      if (leer(f.neu_id)) return null;
      return {
        r, id: f.neu_id, name: leer(f.name) ? "" : String(f.name),
        mit: !e.ohne.some((id) => klein(id) === klein(f.neu_id)),
        rechnet: !!x.berechnet && N(x.berechnet.gueltig) === 1,
        zins: leer(f.kredit_zins) ? "" : f.kredit_zins,
        qng: klein(f.qng) === "ja" ? "ja" : "nein",
        quelle: leer(f.quelle) ? "" : String(f.quelle),
        quote: quoteVon(x), jahr: f.kaufjahr,
        qngGrenze: klein(f.qng_grenze) === "ja" ? "ja" : "nein",
      };
    }).filter(Boolean);
  }

  // ---- Hebel ----
  /** Kandidaten eines Hebels; Kandidat 0 ist immer der Ausgang (Eingaben wie erfasst). */
  function kandidaten(art, x, P, e, hinweise) {
    const f = x.felder, liste = [null];
    if (art === "jahr") {
      if (klein(e.hebelJahr) !== "ja" || !istZahl(f.kaufjahr)) return liste;
      const s = Math.trunc(N(e.jahrSpielraum));
      for (let j = f.kaufjahr - s; j <= f.kaufjahr + s; j++) {
        if (j !== f.kaufjahr && j >= P.Startjahr && j <= P.Endjahr) liste.push(j);
      }
    } else if (art === "qng") {
      if (klein(e.hebel6b) !== "ja" || klein(f.qng) !== "ja") return liste;
      if (RW.konstanten.QUELLEN.every((q) => klein(f[q]) === "")) return liste;
      liste.push(klein(f.qng_grenze) === "ja" ? "nein" : "ja");
    } else {
      if (klein(e.hebelKredit) !== "ja") return liste;
      if (klein(f.kredit_zins) === "") {
        hinweise.push("Kreditquote: Zinssatz fehlt beim Neuobjekt.");
        return liste;
      }
      const schritt = N(e.kreditSchritt) > 0 ? N(e.kreditSchritt) : 0.1;
      for (let q = N(e.kreditMin); q <= N(e.kreditMax) + 0.000001; q += schritt) liste.push(Math.round(q * 1e6) / 1e6);
    }
    return liste;
  }

  /** Eingabezeile eines Neuobjekts mit den gewählten Kandidaten; alle 0 = wie erfasst. */
  function zeileMit(original, x, kand, wahl, felder) {
    if (wahl.every((k) => k === 0)) return original === null || original === undefined ? null : kopie(original);
    let zeile;
    if (original) zeile = kopie(original);
    else {
      // Annahmezeile (reinvestieren = ja): wird zum erfassten Neuobjekt mit denselben Werten
      zeile = {};
      for (const k of felder) zeile[k] = leer(x.felder[k]) ? null : x.felder[k];
      zeile.ak_gub_eingabe = null;
      zeile.ak_geb_eingabe = null;
    }
    ARTEN.forEach((art, i) => {
      const k = wahl[i];
      if (k === 0) return;
      const w = kand[i][k];
      if (art === "jahr") zeile.kaufjahr = w;
      else if (art === "qng") zeile.qng_grenze = w;
      else { zeile.fin_art = "Kredit"; zeile.kredit_betrag = null; zeile.kredit_quote = w; }
    });
    return zeile;
  }

  // ---- Optimieren (modOptimierung.Optimieren) ----
  /**
   * Sucht die beste zulässige Kombination. meta: {fehlerKeys, neuFelder} (Schlüssel der
   * Fehler-Prüfungen und der Eingabefelder eines Neuobjekts). Liefert die optimierte Eingabe
   * (die übergebene bleibt unverändert) und das Protokoll des Laufs.
   */
  function optimiere(eingabe, einst, meta) {
    const e = einstellungen(einst);
    const fehlerKeys = (meta && meta.fehlerKeys) || [];
    const felder = (meta && meta.neuFelder) || [];
    let laeufe = 0;
    const R0 = RW.rechne(eingabe);
    const K0 = kennzahlen(R0, fehlerKeys);
    const fehlerStart = K0.fehler;
    const bewerte = (ein) => {
      const R = RW.rechne(ein);
      laeufe++;
      const K = kennzahlen(R, fehlerKeys);
      const fehler = Math.max(K.fehler - fehlerStart, 0);
      const zulaessig = fehler === 0 && K.kddJahre === 0 && K.liqJahre === 0;
      return { zulaessig, K, R,
        wert: zulaessig ? K.ziel : UNZULAESSIG - fehler * 1e13 - K.kddJahre * 1e10 - K.fehlbetrag };
    };

    // Neuobjekte, die rechnen und optimiert werden sollen
    const zeilen = neuobjektZeilen(R0, e).map((z) => ({ r: z.r, id: z.id, name: z.name,
      vorQuote: z.quote, vorJahr: z.jahr, vor6b: z.qngGrenze, hinweis: "" }));
    const aktive = [];
    for (const z of zeilen) {
      const x = R0.neuobjekte[z.r];
      if (e.ohne.some((id) => klein(id) === klein(z.id))) z.hinweis = "nicht optimiert (optimieren = nein)";
      else if (!x.berechnet || N(x.berechnet.gueltig) !== 1) z.hinweis = "rechnet nicht mit (Status bzw. Kauf aktiv)";
      else {
        const hinweise = [];
        const kand = ARTEN.map((art) => kandidaten(art, x, R0.parameter, e, hinweise));
        z.hinweis = hinweise.join(" ");
        aktive.push({ z, x, kand, wahl: [0, 0, 0], original: (eingabe.neuobjekte || [])[z.r] });
      }
    }

    const mit = (a, wahl) => {
      const ein = Object.assign({}, eingabe, { neuobjekte: (eingabe.neuobjekte || []).slice() });
      for (const b of aktive) {
        const zeile = zeileMit(b.original, b.x, b.kand, b === a ? wahl : b.wahl, felder);
        while (ein.neuobjekte.length <= b.z.r) ein.neuobjekte.push(null);
        ein.neuobjekte[b.z.r] = zeile;
      }
      while (ein.neuobjekte.length && ein.neuobjekte[ein.neuobjekte.length - 1] === null) ein.neuobjekte.pop();
      return ein;
    };

    let best = bewerte(eingabe).wert;
    const runden = Math.max(Math.trunc(N(e.runden)), 1);
    let rundenGelaufen = 0;
    for (let runde = 1; runde <= runden && aktive.length; runde++) {
      rundenGelaufen = runde;
      let verbessert = false;
      for (const a of aktive) {
        ARTEN.forEach((art, i) => {
          if (a.kand[i].length < 2) return;
          const aktuell = a.wahl[i];
          let bestK = aktuell;
          for (let k = 0; k < a.kand[i].length; k++) {
            if (k === aktuell) continue;
            const wahl = a.wahl.slice();
            wahl[i] = k;
            const b = bewerte(mit(a, wahl));
            if (b.wert > best + BESSER_AB) { best = b.wert; bestK = k; }
          }
          if (bestK !== aktuell) { a.wahl[i] = bestK; verbessert = true; }
        });
      }
      if (!verbessert) break;
    }

    const ergebnisEingabe = aktive.length ? mit(null, null) : eingabe;
    const ende = bewerte(ergebnisEingabe);
    const aenderungen = summe(aktive.map((a) => a.wahl.filter((k) => k !== 0).length));
    for (const a of aktive) {
      const geaendert = ARTEN.filter((_, i) => a.wahl[i] !== 0).map((art) => HEBEL_NAME[art]);
      const text = geaendert.length ? "geändert: " + geaendert.join(", ") : "";
      const uebernommen = !a.original && geaendert.length ? "aus „reinvestieren = ja“ als eigenes Neuobjekt übernommen" : "";
      a.z.hinweis = [a.z.hinweis, text, uebernommen].filter(Boolean).join("; ");
    }
    const nach = neuobjektZeilen(ende.R, e);
    for (const z of zeilen) {
      const n = nach.find((y) => y.r === z.r) || {};
      Object.assign(z, { quote: n.quote, jahr: n.jahr, qngGrenze: n.qngGrenze });
    }
    const meldung = !aktive.length ? "Kein Neuobjekt zum Optimieren."
      : !ende.zulaessig ? "Nebenbedingungen mit diesen Hebeln nicht erfüllbar; beste Näherung eingetragen. Ausschüttung senken (Tragbare Ausschüttung ermitteln), Spielraum erweitern oder weitere Käufe erfassen."
      : aenderungen === 0 ? "Eingaben sind bereits die beste gefundene Kombination."
      : "Beste zulässige Kombination eingetragen.";
    return {
      eingabe: ergebnisEingabe, ergebnis: ende.R,
      lauf: { zeit: new Date().toISOString(), meldung, vorher: K0.ziel, nachher: ende.K.ziel,
        liqVorher: K0.liqMin, liqNachher: ende.K.liqMin, erfuellt: ende.zulaessig, aenderungen, laeufe,
        runden: rundenGelaufen, zeilen },
    };
  }

  // ---- Tragbare Ausschüttung (modOptimierung.TragbareAusschuettung) ----
  /** Höchste Ausschüttung p. a. (auf 1.000 € abgerundet), bei der Mindestliquidität und
   *  Mindest-Kapitaldienstdeckung mit den aktuellen Eingaben halten. Der Parameter bleibt. */
  function tragbareAusschuettung(eingabe) {
    let laeufe = 0;
    const haelt = (betrag) => {
      const ein = Object.assign({}, eingabe, { parameter: Object.assign({}, eingabe.parameter, { Ausschuettung: betrag }) });
      laeufe++;
      return kennzahlen(RW.rechne(ein), []).erfuellt;
    };
    let lo = 0;
    if (haelt(0)) {
      let hi = 100000;
      while (haelt(hi)) {
        lo = hi;
        hi *= 2;
        if (hi > 1e9) break;
      }
      while (hi - lo > 1000) {
        const mitte = (lo + hi) / 2;
        if (haelt(mitte)) lo = mitte; else hi = mitte;
      }
    }
    return { betrag: Math.floor(lo / 1000) * 1000, laeufe, zeit: new Date().toISOString() };
  }

  // ---- Varianten (modVarianten) ----
  const tausend = (w) => new Intl.NumberFormat("de-DE", { maximumFractionDigits: 0 }).format(w / 1000);
  function vorgaenge(liste, id, jahr, preis) {
    return liste.filter((x) => !leer(x.felder[id]) && klein(x.felder.aktiv) !== "nein")
      .map((x) => `${String(x.felder[id]).trim()} ${leer(x.felder[jahr]) ? "" : x.felder[jahr]}`
        + (istZahl(x.felder[preis]) ? ` ${tausend(x.felder[preis])} T€` : "")).join("; ");
  }
  /** Kennzahlen einer Variante wie VARIANTEN_SPALTEN, dazu die Eingaben zum Einspielen. */
  function variante(bezeichnung, eingabe, R, pruefungen) {
    const P = R.parameter, vg = R.vergleich[0];
    const K = kennzahlen(R, []);
    const rechnend = R.neuobjekte.filter((x) => x.berechnet && N(x.berechnet.gueltig) === 1);
    return {
      bezeichnung: String(bezeichnung).trim(), datum: new Date().toISOString(),
      A: vg.A, B: vg.B, C: vg.C, Baseline: vg.Baseline, diffB: vg.diffB, diffC: vg.diffC,
      steuer: summe(R.szenarien.A.liq.steuer),
      pruefung: pruefungText(R, pruefungen),
      liqMin: K.liqMin, liqMinJahr: K.liqMinJahr,
      verkaeufe: vorgaenge(R.verkaeufe, "objekt_id", "jahr", "preis"),
      kaeufe: vorgaenge(R.neuobjekte, "neu_id", "kaufjahr", "kaufpreis"),
      steuerwelt: P.Steuerwelt, steuersatzEff: P.SteuersatzEff, mietsteig: P.Mietsteig,
      neuMietsteig: P.NeuMietsteig, erhaltsteig: P.Erhaltsteig, wertsteig: P.Wertsteig,
      alternativrendite: P.Alternativrendite, ausschuettung: N(P.Ausschuettung),
      kreditquoten: rechnend.filter((x) => N(x.berechnet.ak_gesamt) > 0)
        .map((x) => `${String(x.felder.neu_id).trim()} ${Math.round((N(x.berechnet.kredit) / x.berechnet.ak_gesamt) * 100)} %`).join("; "),
      qngGrenze: rechnend.filter((x) => klein(x.felder.qng_grenze) === "ja").map((x) => String(x.felder.neu_id).trim()).join("; "),
      eingabe: kopie(eingabe),
    };
  }
  /** Spalten der Variantentabelle: [Schlüssel, Überschrift, Format]. */
  const VARIANTEN_SPALTEN = [
    ["A", "Endvermögen A", "euro"], ["B", "Endvermögen B", "euro"], ["C", "Endvermögen C", "euro"],
    ["Baseline", "Endvermögen Halten", "euro"], ["diffB", "A − B", "euro"], ["diffC", "A − C", "euro"],
    ["steuer", "Steuer A gesamt", "euro"], ["pruefung", "Prüfung", "text"],
    ["liqMin", "Liquidität tiefster Stand A", "euro"], ["liqMinJahr", "im Jahr", "jahr"],
    ["verkaeufe", "Verkäufe (ObjektID Jahr Preis)", "text"], ["kaeufe", "Käufe (NeuID Jahr Preis)", "text"],
    ["steuerwelt", "Steuerwelt", "text"], ["steuersatzEff", "angewandter Steuersatz", "prozent"],
    ["mietsteig", "Mietsteigerung", "prozent"], ["neuMietsteig", "Mietsteigerung Neuobjekte", "prozent"],
    ["erhaltsteig", "Erhaltungssteigerung", "prozent"], ["wertsteig", "Wertsteigerung", "prozent"],
    ["alternativrendite", "Alternativrendite", "prozent"], ["ausschuettung", "Ausschüttung p. a.", "euro"],
    ["kreditquoten", "Kreditquote je Kauf", "text"], ["qngGrenze", "§ 6b nur bis QNG-Grenze", "text"],
  ];

  return { EINSTELLUNGEN, einstellungen, kennzahlen, pruefungText, neuobjektZeilen, optimiere,
    tragbareAusschuettung, variante, VARIANTEN_SPALTEN, MAX_VARIANTEN, HEBEL_NAME };
});
