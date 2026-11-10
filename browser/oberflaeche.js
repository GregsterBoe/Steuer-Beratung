/*
 * Oberfläche der Browser-Fassung: Eingaben, Ergebnisse, Rechenweg je Zahl.
 * Rechnet bei jeder Änderung mit Rechenwerk.rechne neu; Daten bleiben im Browser.
 */
(function () {
  "use strict";
  const M = window.PROGNOSE_META;
  const START = window.PROGNOSE_START;
  const RW = window.Rechenwerk;
  const WEG = window.Rechenweg;
  const SPEICHER = "prognosemodell-entwurf";
  const OHNE_BROWSER = ["bwa_zuordnung", "kst_abweichung", "bruecke", "cashflow"];
  const SZ_FARBE = { A: "var(--a)", B: "var(--b)", C: "var(--c)", Baseline: "var(--baseline)" };
  const SZ_NAME = { A: "§ 6b-Kette (Plan)", B: "sofort versteuern, anlegen", C: "sofort versteuern, kaufen", Baseline: "alles halten" };
  const $ = (sel) => document.querySelector(sel);
  const leer = (v) => v === null || v === undefined || v === "";
  const istZahl = (v) => typeof v === "number" && isFinite(v);
  const kopie = (x) => JSON.parse(JSON.stringify(x));

  const zustand = {
    eingabe: null, ergebnis: null, fehler: null, ansicht: "ueberblick", sz: "A",
    prgObjekt: "b:0", weg: [], gewaehlt: null, quelle: "datei",
  };
  const startKennung = String(hash(JSON.stringify(START) + (window.PROGNOSE_ERZEUGT || "")));

  // ---------------------------------------------------------------- Formate
  const nf = (min, max) => new Intl.NumberFormat("de-DE", { minimumFractionDigits: min, maximumFractionDigits: max });
  const F0 = nf(0, 0), F1 = nf(1, 1), F2 = nf(2, 2);
  function format(wert, art) {
    if (leer(wert)) return "";
    if (typeof wert === "string") return wert;
    if (typeof wert === "boolean") return wert ? "ja" : "nein";
    if (!istZahl(wert)) return String(wert);
    switch (art) {
      case "euro": return F0.format(wert) + " €";
      case "euro2": return F2.format(wert) + " €";
      case "prozent": return F2.format(wert * 100) + " %";
      case "prozent1": return F1.format(wert * 100) + " %";
      case "prozent0": return F0.format(wert * 100) + " %";
      case "jahr": return String(Math.round(wert));
      case "faktor": return F1.format(wert) + "-fach";
      case "dez1": return F1.format(wert);
      case "dez2": return F2.format(wert);
      case "zahl": return F0.format(wert);
      default: return Math.abs(wert) >= 100 ? F0.format(wert) : String(wert);
    }
  }
  /** Wert zum Bearbeiten: ohne Einheit, Prozent in Prozentpunkten. */
  function formatEingabe(wert, art) {
    if (leer(wert)) return "";
    if (!istZahl(wert)) return String(wert);
    if (art === "prozent" || art === "prozent1" || art === "prozent0") return String(+(wert * 100).toFixed(6)).replace(".", ",");
    if (art === "jahr") return String(wert);
    if (art === "euro" || art === "zahl") return F0.format(wert) === "0" && wert !== 0 ? String(wert).replace(".", ",") : nf(0, 2).format(wert);
    return String(wert).replace(".", ",");
  }
  function lese(text, art) {
    const t = String(text).trim();
    if (t === "") return null;
    if (art === "text" || art === undefined) return t;
    let s = t.replace(/[€%\s]/g, "").replace(/−/g, "-");
    if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
    else if ((s.match(/\./g) || []).length > 1 || /^\-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, "");
    const z = Number(s);
    if (!isFinite(z)) return t;
    if (art === "prozent" || art === "prozent1" || art === "prozent0") return z / 100;
    if (art === "jahr") return Math.round(z);
    return z;
  }
  function artVon(pfad) {
    const t = pfad.split(":");
    const such = (liste, key) => { const x = (liste || []).find((s) => s.key === key || s.name === key); return x ? x.format : null; };
    switch (t[0]) {
      case "par": return such(M.parameter, t[1]) || "zahl";
      case "obj": return such(M.felder.objekte, t[2]) || "euro";
      case "prg": return such(M.spalten.prognose, t[3]) || "euro";
      case "vk": return such(M.spalten.verkauf, t[2]) || such(M.felder.verkaeufe, t[2]) || "euro";
      case "rl": return such(M.spalten.ruecklage, t[2]) || "euro";
      case "rls": return such(M.spalten.spiegel, t[1]) || "euro";
      case "ne": return such(M.spalten.neu, t[2]) || such(M.felder.neuobjekte, t[2]) || "euro";
      case "dl": return t.length > 4 ? "euro" : such(M.spalten.darlehen, t[3]) || "euro";
      case "liq": return such(M.spalten.liq, t[2]) || "euro";
      case "aus": return such(M.spalten.aus, t[2]) || "euro";
      default: return "euro";
    }
  }
  const esc = (s) => String(s === null || s === undefined ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  function hash(s) { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return h; }
  function kurz(wert) {
    if (!istZahl(wert)) return "";
    const a = Math.abs(wert);
    if (a >= 1e6) return F1.format(wert / 1e6) + " Mio";
    if (a >= 1e3) return F0.format(wert / 1e3) + " T";
    return F0.format(wert);
  }

  // ---------------------------------------------------------------- Zustand und Rechnung
  function neuRechnen() {
    const t0 = performance.now();
    try {
      zustand.ergebnis = RW.rechne(zustand.eingabe);
      zustand.fehler = null;
    } catch (e) {
      zustand.fehler = e;
      console.error(e);
    }
    zustand.dauer = performance.now() - t0;
    speichern();
    zeichnen();
  }
  function speichern() {
    try { localStorage.setItem(SPEICHER, JSON.stringify({ start: startKennung, eingabe: zustand.eingabe })); } catch (e) { /* ohne Speicher */ }
  }
  function laden() {
    try {
      const roh = localStorage.getItem(SPEICHER);
      if (!roh) return null;
      const s = JSON.parse(roh);
      return s && s.start === startKennung ? s.eingabe : null;
    } catch (e) { return null; }
  }
  function setzeEingabe(neu, quelle) {
    zustand.eingabe = neu;
    zustand.quelle = quelle;
    zustand.weg = [];
    neuRechnen();
  }

  // ---------------------------------------------------------------- Navigation
  const ANSICHTEN = [
    ["Ergebnis", [["ueberblick", "Überblick"], ["vergleich", "Szenarienvergleich"], ["liquiditaet", "Liquidität und Steuer"], ["vermoegen", "Vermögen"]]],
    ["Eingaben", [["parameter", "Parameter"], ["objekte", "Objekte"], ["verkaeufe", "Verkäufe"], ["neuobjekte", "Neuobjekte"], ["daten", "Anlagen und Planwerte"]]],
    ["Rechenweg", [["prognose", "Prognose je Objekt"], ["ruecklagen", "§ 6b-Rücklagen"], ["darlehen", "Darlehen"]]],
    ["Kontrolle", [["pruefung", "Prüfung"]]],
  ];
  function zeichneNavi() {
    const R = zustand.ergebnis;
    const anzahl = R ? {
      objekte: R.objekte.filter(Boolean).length, verkaeufe: R.verkaeufe.length,
      neuobjekte: R.neuobjekte.filter((x) => x.berechnet).length,
    } : {};
    $("#navi").innerHTML = ANSICHTEN.map(([gruppe, liste]) => `<div class="gruppe">${gruppe}</div>` + liste.map(([id, titel]) =>
      `<button data-ansicht="${id}" class="${zustand.ansicht === id ? "aktiv" : ""}">${titel}${anzahl[id] !== undefined ? `<span class="anzahl">${anzahl[id]}</span>` : ""}</button>`).join("")).join("");
  }

  // ---------------------------------------------------------------- Bausteine
  function zelle(pfad, extra = "") {
    const R = zustand.ergebnis;
    const w = WEG.wertVon(pfad, R);
    const art = artVon(pfad);
    const neg = istZahl(w) && w < -0.005 && art.startsWith("euro") ? " neg" : "";
    const gew = zustand.gewaehlt === pfad ? " gewaehlt" : "";
    return `<td class="z klickbar${neg}${gew} ${extra}" data-pfad="${pfad}">${esc(format(w, art))}</td>`;
  }
  function jahresKopf(jahre, summe, erste = "") {
    return `<thead><tr><th class="fest">${erste}</th>${jahre.map((t) => `<th class="z">${t}</th>`).join("")}${summe ? '<th class="z">Summe</th>' : ""}</tr></thead>`;
  }
  /** Tabelle Zeilen × Jahre. zeilen: [{label, pfad(i), summe, klasse, titel}] oder {abschnitt}. */
  function jahresTabelle(zeilen, mitSumme = false) {
    const R = zustand.ergebnis;
    const jahre = R.jahre;
    const koerper = zeilen.map((z) => {
      if (z.abschnitt) return `<tr class="abschnitt"><td class="fest">${esc(z.abschnitt)}</td><td colspan="${jahre.length + (mitSumme ? 1 : 0)}"></td></tr>`;
      const werte = jahre.map((_, i) => WEG.wertVon(z.pfad(i), R));
      const summe = mitSumme && z.summe !== false ? werte.reduce((a, b) => a + (istZahl(b) ? b : 0), 0) : null;
      return `<tr class="${z.klasse || ""}"><td class="fest"${z.titel ? ` title="${esc(z.titel)}"` : ""}>${esc(z.label)}</td>${jahre.map((_, i) => zelle(z.pfad(i))).join("")}${mitSumme ? `<td class="z">${summe === null ? "" : esc(format(summe, artVon(z.pfad(0))))}</td>` : ""}</tr>`;
    }).join("");
    return `<div class="tabelle-rahmen"><table>${jahresKopf(jahre, mitSumme)}<tbody>${koerper}</tbody></table></div>`;
  }
  function reiter(aktiv, attribut = "data-sz") {
    return `<div class="reiter">${M.szenarien.map((s) => `<button ${attribut}="${s.key}" class="${aktiv === s.key ? "aktiv" : ""}">${esc(s.kurz)}</button>`).join("")}</div>`;
  }
  function seitenkopf(titel, text, rechts = "") {
    return `<div class="seitenkopf"><div><h2>${esc(titel)}</h2>${text ? `<p>${text}</p>` : ""}</div>${rechts}</div>`;
  }
  function statusKlasse(s) {
    if (s === RW.konstanten.STATUS_OK) return "ok";
    if (s === RW.konstanten.STATUS_AUS) return "aus";
    return s ? "fehler" : "";
  }

  // ---------------------------------------------------------------- Diagramm
  const DIAGRAMME = {};
  function linienDiagramm(id, serien, jahre, optionen = {}) {
    const B = 640, H = optionen.hoehe || 250, L = 62, Rr = 12, O = 12, U = 26;
    const alle = serien.flatMap((s) => s.werte).concat(optionen.grenze !== undefined ? [optionen.grenze] : []).filter(istZahl);
    let min = Math.min(0, ...alle), max = Math.max(0, ...alle);
    if (max === min) max = min + 1;
    const schritt = schoeneSchritte(min, max, 5);
    min = Math.floor(min / schritt) * schritt;
    max = Math.ceil(max / schritt) * schritt;
    const x = (i) => L + (i * (B - L - Rr)) / Math.max(jahre.length - 1, 1);
    const y = (w) => O + ((max - w) * (H - O - U)) / (max - min);
    let svg = `<svg class="diagramm" viewBox="0 0 ${B} ${H}" data-diagramm="${id}" role="img" aria-label="${esc(optionen.titel || "")}">`;
    for (let w = min; w <= max + schritt / 2; w += schritt) {
      svg += `<line class="${Math.abs(w) < schritt / 1e6 ? "nulllinie" : "gitter"}" x1="${L}" x2="${B - Rr}" y1="${y(w)}" y2="${y(w)}"/>`;
      svg += `<text x="${L - 8}" y="${y(w) + 4}" text-anchor="end">${kurz(w)}</text>`;
    }
    const jeder = Math.ceil(jahre.length / 10);
    const letzte = jahre.length - 1;
    jahre.forEach((t, i) => {
      const zeigen = i === letzte || (i % jeder === 0 && letzte - i >= jeder);
      if (zeigen) svg += `<text x="${x(i)}" y="${H - 8}" text-anchor="${i === letzte ? "end" : "middle"}">${t}</text>`;
    });
    if (optionen.grenze !== undefined && optionen.grenze !== 0) {
      svg += `<line class="fuehrung" x1="${L}" x2="${B - Rr}" y1="${y(optionen.grenze)}" y2="${y(optionen.grenze)}"/>`;
    }
    serien.forEach((s) => {
      const d = s.werte.map((w, i) => `${i ? "L" : "M"}${x(i).toFixed(1)},${y(istZahl(w) ? w : 0).toFixed(1)}`).join("");
      svg += `<path d="${d}" fill="none" stroke="${s.farbe}" stroke-width="${s.dick ? 2.5 : 2}" stroke-linejoin="round" stroke-linecap="round"${s.gestrichelt ? ' stroke-dasharray="5 4"' : ""}/>`;
    });
    svg += `<line class="fuehrung" data-fuehrung x1="0" x2="0" y1="${O}" y2="${H - U}" visibility="hidden"/>`;
    svg += serien.map((s, k) => `<circle data-punkt="${k}" r="4" fill="${s.farbe}" stroke="var(--flaeche)" stroke-width="2" visibility="hidden"/>`).join("");
    svg += `<rect x="${L}" y="${O}" width="${B - L - Rr}" height="${H - O - U}" fill="transparent" data-flaeche/>`;
    svg += "</svg>";
    DIAGRAMME[id] = { serien, jahre, x, y, B, L, Rr };
    return svg;
  }
  function schoeneSchritte(min, max, n) {
    const roh = (max - min) / n;
    const p = Math.pow(10, Math.floor(Math.log10(roh)));
    return [1, 2, 2.5, 5, 10].map((f) => f * p).find((s) => s >= roh) || roh;
  }
  function diagrammHover(e) {
    const svg = e.target.closest("svg[data-diagramm]");
    const tip = $("#tooltip");
    if (!svg) { tip.hidden = true; return; }
    const d = DIAGRAMME[svg.dataset.diagramm];
    if (!d) return;
    const r = svg.getBoundingClientRect();
    const px = ((e.clientX - r.left) / r.width) * d.B;
    const i = Math.max(0, Math.min(d.jahre.length - 1, Math.round(((px - d.L) / (d.B - d.L - d.Rr)) * (d.jahre.length - 1))));
    const f = svg.querySelector("[data-fuehrung]");
    f.setAttribute("x1", d.x(i)); f.setAttribute("x2", d.x(i)); f.setAttribute("visibility", "visible");
    svg.querySelectorAll("[data-punkt]").forEach((c) => {
      const s = d.serien[Number(c.dataset.punkt)];
      c.setAttribute("cx", d.x(i)); c.setAttribute("cy", d.y(istZahl(s.werte[i]) ? s.werte[i] : 0)); c.setAttribute("visibility", "visible");
    });
    tip.innerHTML = `<b>${d.jahre[i]}</b>` + d.serien.map((s) => `<div class="tt-zeile"><span><i class="punkt" style="background:${s.farbe}"></i>${esc(s.name)}</span><span class="zahl">${esc(format(s.werte[i], "euro"))}</span></div>`).join("");
    tip.hidden = false;
    const tx = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 8);
    tip.style.left = tx + "px";
    tip.style.top = (e.clientY + 14) + "px";
    svg.dataset.index = i;
  }
  function diagrammWeg(e) {
    const svg = e.target.closest("svg[data-diagramm]");
    if (!svg) return;
    $("#tooltip").hidden = true;
    svg.querySelectorAll("[data-punkt],[data-fuehrung]").forEach((c) => c.setAttribute("visibility", "hidden"));
  }

  // ---------------------------------------------------------------- Ansichten: Ergebnis
  function ansichtUeberblick() {
    const R = zustand.ergebnis, P = R.parameter;
    const ende = R.jahre[R.jahre.length - 1];
    const vg = R.vergleich[0];
    const reihen = M.szenarien.map((s) => ({ key: s.key, wert: vg[s.key] }));
    const beste = reihen.reduce((a, b) => (b.wert > a.wert ? b : a));
    const vorsprung = beste.wert - vg.Baseline;
    const empf = beste.key === "Baseline"
      ? `<strong>Halten ist vorn.</strong><p>Alles halten ergibt ${ende} das höchste Endvermögen nach latenter Steuer: ${esc(format(vg.Baseline, "euro"))}. Der Plan liegt ${esc(format(vg.Baseline - vg.A, "euro"))} darunter.</p>`
      : `<strong>${esc(SZ_NAME[beste.key])}: höchstes Endvermögen ${ende}</strong><p>${esc(format(beste.wert, "euro"))} nach latenter Steuer, ${esc(format(vorsprung, "euro"))} mehr als alles halten.${beste.key !== "A" ? ` Der Plan (§ 6b-Kette) liegt ${esc(format(beste.wert - vg.A, "euro"))} darunter.` : ""}</p>`;
    const L = R.szenarien.A.liq;
    let minI = 0;
    L.kum.forEach((w, i) => { if (w < L.kum[minI]) minI = i; });
    const steuerA = L.steuer.reduce((a, b) => a + b, 0);
    const ausschuettung = L.ausschuettung.reduce((a, b) => a + b, 0);
    const pr = pruefungGesamt(R);
    const karten = M.szenarien.map((s) => `<div class="karte${s.key === beste.key ? " beste" : ""}">
        <div class="etikett"><i class="punkt" style="background:${SZ_FARBE[s.key]}"></i>${esc(s.kurz)}</div>
        <div class="gross klickbar" data-pfad="vg:0:${s.key}">${esc(format(vg[s.key], "euro"))}</div>
        <div class="unter">${s.key === "Baseline" ? "Endvermögen nach latenter Steuer" : `${vg[s.key] - vg.Baseline >= 0 ? "+" : ""}${esc(format(vg[s.key] - vg.Baseline, "euro"))} gegen Halten`}</div>
      </div>`).join("");
    const serien = M.szenarien.map((s) => ({ name: s.kurz, farbe: SZ_FARBE[s.key], werte: R.szenarien[s.key].aus.vermoegen_netto, dick: s.key === "A" }));
    const liqSerien = [
      { name: "Plan (A)", farbe: SZ_FARBE.A, werte: L.kum, dick: true },
      { name: "alles halten", farbe: SZ_FARBE.Baseline, werte: R.szenarien.Baseline.liq.kum },
    ];
    return seitenkopf("Überblick", `Prognose ${R.jahre[0]} bis ${ende}, Basisjahr ${P.Basisjahr}, ${esc(P.Steuerwelt)}, angewandter Steuersatz ${esc(format(P.SteuersatzEff, "prozent"))}. Jede Zahl lässt sich anklicken: rechts erscheint ihr Rechenweg.`)
      + `<div class="empfehlung">${empf}</div>`
      + `<div class="karten">${karten}</div>`
      + `<div class="zwei">
          <div class="panel"><h3>Gesamtvermögen nach latenter Steuer</h3>
            <div class="legende">${serien.map((s) => `<span><i class="punkt" style="background:${s.farbe}"></i>${esc(s.name)}</span>`).join("")}</div>
            ${linienDiagramm("vermoegen", serien, R.jahre, { titel: "Gesamtvermögen" })}</div>
          <div class="panel"><h3>Liquidität kumuliert</h3>
            <div class="legende">${liqSerien.map((s) => `<span><i class="punkt" style="background:${s.farbe}"></i>${esc(s.name)}</span>`).join("")}${istZahl(P.MindestLiq) && P.MindestLiq !== 0 ? `<span>– – Mindestliquidität ${esc(format(P.MindestLiq, "euro"))}</span>` : ""}</div>
            ${linienDiagramm("liquiditaet", liqSerien, R.jahre, { titel: "Liquidität", grenze: istZahl(P.MindestLiq) ? P.MindestLiq : 0 })}</div>
        </div>`
      + `<h3>Plan (A) auf einen Blick</h3><div class="karten">
          <div class="karte"><div class="etikett">tiefste Liquidität</div><div class="gross klickbar" data-pfad="liq:A:kum:${minI}">${esc(format(L.kum[minI], "euro"))}</div><div class="unter">im Jahr ${R.jahre[minI]}</div></div>
          <div class="karte"><div class="etikett">Steuer über alle Jahre</div><div class="gross klickbar" data-pfad="vg:14:A">${esc(format(steuerA, "euro"))}</div><div class="unter">Halten: ${esc(format(R.vergleich[14].Baseline, "euro"))}</div></div>
          <div class="karte"><div class="etikett">Ausschüttungen</div><div class="gross klickbar" data-pfad="vg:19:A">${esc(format(ausschuettung, "euro"))}</div><div class="unter">${istZahl(P.Ausschuettung) && P.Ausschuettung ? esc(format(P.Ausschuettung, "euro")) + " p. a." : "keine"}</div></div>
          <div class="karte"><div class="etikett">Restschuld Darlehen ${ende}</div><div class="gross klickbar" data-pfad="vg:17:A">${esc(format(R.vergleich[17].A, "euro"))}</div><div class="unter">Zinsen gesamt ${esc(format(R.vergleich[18].A, "euro"))}</div></div>
          <div class="karte"><div class="etikett">Plausibilitätsprüfung</div><div class="gross ${pr.klasse}" style="color:var(--${pr.klasse === "ok" ? "gut" : pr.klasse === "fehler" ? "schlecht" : "warnung"})">${esc(pr.text)}</div><div class="unter"><a href="#" data-ansicht="pruefung">Einzelheiten</a></div></div>
        </div>`;
  }

  function ansichtVergleich() {
    const R = zustand.ergebnis;
    const kopf = `<thead><tr><th class="fest">Kennzahl</th>${M.szenarien.map((s) => `<th class="z"><i class="punkt" style="background:${SZ_FARBE[s.key]}"></i> ${esc(s.kurz)}</th>`).join("")}<th class="z">A − B</th><th class="z">A − C</th><th class="z">A − Halten</th></tr></thead>`;
    const zeilen = R.vergleich.map((z, k) => `<tr${k === 0 ? ' class="summe"' : ""}><td class="fest" title="${esc(M.vergleich[k].erlaeuterung)}">${esc(z.kennzahl)}</td>${M.szenarien.map((s) => zelle(`vg:${k}:${s.key}`)).join("")}${["diffB", "diffC", "diffBaseline"].map((d) => `<td class="z${z[d] < -0.005 ? " neg" : ""}">${esc(format(z[d], "euro"))}</td>`).join("")}</tr>`).join("");
    return seitenkopf("Szenarienvergleich", `Kennzahlen am Ende des Rasters (${R.jahre[R.jahre.length - 1]}) bzw. Summe über alle Jahre. A Plan: § 6b-Kette wie erfasst. B: jeder Gewinn sofort versteuert, Neuobjekte mit Rücklage entfallen, das Geld bleibt angelegt. C: sofort versteuert, Neuobjekte trotzdem gekauft. Halten: nichts verkaufen, nichts kaufen. A − C zeigt die reine Wirkung von § 6b.`)
      + `<div class="tabelle-rahmen"><table>${kopf}<tbody>${zeilen}</tbody></table></div>`;
  }

  const LIQ_GRUPPEN = [
    ["Laufendes Ergebnis", ["einnahmen", "ausgaben", "afa", "ergebnis"]],
    ["Steuer", ["verkauf", "zins", "kreditzins", "zve", "vortrag_genutzt", "bemessung", "vortrag", "steuer"]],
    ["Geldfluss", ["verkaufserloes", "rueckfluss", "kauf", "kredit", "tilgung", "abloesung", "ausschuettung", "zufluss", "kum"]],
    ["Darlehen", ["restschuld"]],
  ];
  const BESTAENDE = ["vortrag", "kum", "restschuld", "steuer_kum", "verkehrswert", "buchwert", "stille_reserven", "ruecklage", "liquiditaet", "vermoegen", "latente_steuer", "vermoegen_netto"];
  const UNTER = ["rueckfluss", "abloesung", "vortrag_genutzt"];
  const SUMMENZEILEN = ["ergebnis", "steuer", "zufluss", "kum", "vermoegen_netto", "nach_steuer"];
  function tabelleGruppen(tab, gruppen) {
    const sp = M.spalten[tab];
    const zeilen = [];
    gruppen.forEach(([titel, keys]) => {
      zeilen.push({ abschnitt: titel });
      keys.forEach((k) => {
        const s = sp.find((x) => x.key === k);
        zeilen.push({ label: s ? s.ueberschrift : k, pfad: (i) => `${tab}:${zustand.sz}:${k}:${i}`, summe: !BESTAENDE.includes(k),
          klasse: SUMMENZEILEN.includes(k) ? "summe" : UNTER.includes(k) ? "unter" : "" });
      });
    });
    return jahresTabelle(zeilen, true);
  }
  function ansichtLiquiditaet() {
    const s = M.szenarien.find((x) => x.key === zustand.sz);
    return seitenkopf("Liquidität und Steuer", `${esc(s.titel)}. Steuer mit Verlustvortrag; die Liquidität wird mit der Rendite der Alternativanlage verzinst.`)
      + reiter(zustand.sz) + tabelleGruppen("liq", LIQ_GRUPPEN);
  }
  const AUS_GRUPPEN = [
    ["Ergebnis", ["ergebnis", "verkauf", "zins", "kreditzins", "guv", "steuer", "nach_steuer", "steuer_kum"]],
    ["Vermögen am Jahresende", ["verkehrswert", "buchwert", "stille_reserven", "ruecklage", "vortrag", "liquiditaet", "restschuld", "vermoegen", "latente_steuer", "vermoegen_netto"]],
  ];
  function ansichtVermoegen() {
    const s = M.szenarien.find((x) => x.key === zustand.sz);
    return seitenkopf("Vermögen", `${esc(s.titel)}. Gesamtvermögen = Verkehrswert + Liquidität − Restschuld; die latente Steuer fiele an, wenn alle Objekte zum Verkehrswert verkauft und die Rücklage aufgelöst würden.`)
      + reiter(zustand.sz) + tabelleGruppen("aus", AUS_GRUPPEN);
  }

  // ---------------------------------------------------------------- Ansichten: Eingaben
  function eingabeFeld(ziel, art, roh, wirksam, herkunft, feld, inTabelle = true) {
    const auswahl = feld && feld.auswahl;
    let klasse = "eingabe";
    if (leer(roh) && (herkunft === "annahme" || herkunft === "abgeleitet")) klasse += " annahme";
    if (herkunft === "kostenstelle" || herkunft === "anlagen") klasse += " eingelesen";
    if (feld && feld.pflicht && leer(roh) && leer(wirksam)) klasse += " pflicht-fehlt";
    const titel = feld && feld.hinweis ? ` title="${esc(feld.hinweis)}"` : "";
    let html;
    if (auswahl) {
      const opts = ["", ...auswahl].map((o) => `<option value="${esc(o)}"${String(roh ?? "") === String(o) ? " selected" : ""}>${o === "" ? (leer(wirksam) || herkunft !== "annahme" ? "" : "(" + esc(wirksam) + ")") : esc(o)}</option>`).join("");
      html = `<select data-ziel="${ziel}" data-art="text"${titel}>${opts}</select>`;
    } else {
      const zahl = art !== "text" ? " z" : "";
      const platz = leer(roh) && !leer(wirksam) ? ` placeholder="${esc(formatEingabe(wirksam, art))}"` : "";
      html = `<input class="${zahl}" data-ziel="${ziel}" data-art="${art}" value="${esc(formatEingabe(roh, art))}"${platz}${titel}>`;
    }
    return inTabelle ? `<td class="${klasse}">${html}</td>` : { klasse, html };
  }
  function formFeld(ziel, feld, roh, wirksam, herkunft, label) {
    const e = eingabeFeld(ziel, feld.format, roh, wirksam, herkunft, feld, false);
    const einheit = feld.format.startsWith("prozent") ? " in %" : "";
    return `<div class="feld ${e.klasse.includes("annahme") ? "annahme" : ""}"><label title="${esc(feld.hinweis || "")}">${esc(label || feld.ueberschrift)}${einheit}${feld.pflicht ? ' <span class="pflicht">*</span>' : ""}</label>${e.html}</div>`;
  }

  function ansichtParameter() {
    const R = zustand.ergebnis;
    const gruppen = [];
    let aktuell = { titel: "Allgemein", liste: [] };
    gruppen.push(aktuell);
    M.parameter.forEach((p) => {
      if (p.abschnitt) { aktuell = { titel: p.abschnitt, liste: [] }; gruppen.push(aktuell); }
      aktuell.liste.push(p);
    });
    const html = gruppen.map((g, n) => `<details class="feldgruppe"${n < 2 ? " open" : ""}><summary>${esc(g.titel)}<span class="zusatz">${g.liste.length} Parameter</span></summary><div><div class="formular">${g.liste.map((p) => {
      if (p.berechnet) return `<div class="feld"><label>${esc(p.bezeichnung)}</label><div class="wert">${esc(format(R.parameter[p.name], p.format))}</div><div class="hilfe">${esc(p.erlaeuterung)}</div></div>`;
      const feld = { format: p.format, auswahl: p.auswahl, hinweis: p.erlaeuterung, ueberschrift: p.bezeichnung };
      const roh = zustand.eingabe.parameter[p.name];
      const e = eingabeFeld(`par:${p.name}`, p.format, roh, null, "", feld, false);
      const abweichend = !leer(p.wert) && roh !== p.wert ? ` <span class="leise">(Vorgabe ${esc(format(p.wert, p.format))})</span>` : "";
      return `<div class="feld"><label>${esc(p.bezeichnung)}${p.format.startsWith("prozent") ? " in %" : ""}${abweichend}</label>${e.html}<div class="hilfe">${esc(p.erlaeuterung)}</div></div>`;
    }).join("")}</div></div></details>`).join("");
    return seitenkopf("Parameter", "Alle Annahmen der Rechnung. Leeren setzt die Vorgabe wieder ein. Sätze sind Platzhalter und vor dem Echteinsatz fachlich zu prüfen.") + html;
  }

  const FARBLEGENDE = `<div class="farblegende"><span><i style="background:var(--eingabe-hg)"></i>Eingabe</span><span><i style="background:var(--annahme-hg)"></i>Annahme (leer gelassen, Wert violett)</span><span><i style="background:var(--eingelesen-hg)"></i>aus Buchhaltung oder Anlagen</span><span><i style="box-shadow:inset 0 0 0 1px var(--schlecht)"></i>Pflichtfeld fehlt</span></div>`;

  const OBJ_GRUPPEN = [
    ["Laufende Werte im Basisjahr", ["objekt_id", "name", "art", "miete", "erhaltung", "weitere_einnahmen", "weitere_ausgaben", "afa_bwa", "zinsen"]],
    ["Fortschreibung", ["mietsteig", "erhaltsteig", "miete_fix_ab"]],
    ["Steuerliche Stammdaten", ["verkehrswert", "vk_quote_gebaeude", "afa_satz", "kaufjahr", "ak_gebaeude", "ak_gub", "restbuchwert", "afa_jahr"]],
    ["Erhaltung nach Alter", ["baujahr", "san_jahr", "san_betrag"]],
    ["Darlehen", ["restschuld", "zinssatz", "rate"]],
  ];
  function ansichtObjekte() {
    const R = zustand.ergebnis;
    const felder = Object.fromEntries(M.felder.objekte.map((f) => [f.key, f]));
    const kopf1 = `<tr class="obergruppe"><th class="fest"></th>${OBJ_GRUPPEN.map(([t, k]) => `<th colspan="${k.length - (t.startsWith("Laufende") ? 1 : 0)}">${esc(t)}</th>`).join("")}<th></th><th></th></tr>`;
    const kopf2 = `<tr>${OBJ_GRUPPEN.flatMap(([, k]) => k).map((key, n) => { const f = felder[key]; return `<th class="${n === 0 ? "fest" : ""}${f.format !== "text" ? " z" : ""}" title="${esc(f.hinweis || "")}">${esc(f.ueberschrift)}${f.pflicht ? ' <span class="pflicht">*</span>' : ""}${f.format.startsWith("prozent") ? " %" : ""}</th>`; }).join("")}<th>Status</th><th></th></tr>`;
    const zeilen = zustand.eingabe.objekte.map((roh, j) => {
      const o = R.objekte[j];
      const zellen = OBJ_GRUPPEN.flatMap(([, k]) => k).map((key, n) => {
        const f = felder[key];
        const w = o ? o.felder[key] : "";
        const h = o ? o.herkunft[key] : "";
        const td = eingabeFeld(`obj:${j}:${key}`, f.format, roh[key], w, roh.herkunft && roh.herkunft[key] ? roh.herkunft[key] : h, f);
        return n === 0 ? td.replace('<td class="', '<td class="fest ') : td;
      }).join("");
      const st = o ? o.status : "";
      return `<tr>${zellen}<td class="status ${statusKlasse(st)}">${esc(st)}</td><td><button class="knopf klein gefahr" data-aktion="obj-weg" data-nr="${j}" title="Objekt entfernen">✕</button></td></tr>`;
    }).join("");
    return seitenkopf("Objekte", "Bestandsobjekte und Kostenstellen. Leere Felder rechnen mit der Annahme vom Parameterblatt (violett, kursiv); ein eingetippter Wert ersetzt sie, Leeren stellt sie wieder her. Wert anklicken in den Rechenblättern zeigt, woher er kommt.",
      `<button class="knopf haupt" data-aktion="obj-neu">+ Objekt</button>`)
      + FARBLEGENDE
      + (zeilen ? `<div class="tabelle-rahmen"><table><thead>${kopf1}${kopf2}</thead><tbody>${zeilen}</tbody></table></div>` : `<div class="leer-hinweis">Noch keine Objekte. „+ Objekt“ legt eins an; Pflicht sind ObjektID und Miete.</div>`);
  }

  const VK_FELDER = ["objekt_id", "weitere_ids", "aktiv", "jahr", "preis", "kosten", "anteil_gub", "gub_betrag", "nutzung_6b", "neubau_6b", "reinvest"];
  const VK_ERGEBNIS = ["nettoerloes", "buchwert_geb", "gewinn_geb", "gewinn_gub", "gewinn", "vorbesitz"];
  function ansichtVerkaeufe() {
    const R = zustand.ergebnis;
    const felder = Object.fromEntries(M.felder.verkaeufe.map((f) => [f.key, f]));
    const sp = Object.fromEntries(M.spalten.verkauf.map((s) => [s.key, s]));
    const ids = R.objekte.filter(Boolean).map((o) => o.felder.objekt_id);
    const kopf = `<thead><tr class="obergruppe"><th class="fest"></th><th colspan="${VK_FELDER.length - 1}">Eingaben</th><th colspan="${VK_ERGEBNIS.length + 1}">Ergebnis</th><th></th><th></th></tr><tr>${VK_FELDER.map((k, n) => `<th class="${n === 0 ? "fest" : ""}${felder[k].format !== "text" ? " z" : ""}" title="${esc(felder[k].hinweis || "")}">${esc(felder[k].ueberschrift)}${felder[k].pflicht ? ' <span class="pflicht">*</span>' : ""}${felder[k].format.startsWith("prozent") ? " %" : ""}</th>`).join("")}${VK_ERGEBNIS.map((k) => `<th class="z">${esc(sp[k].ueberschrift)}</th>`).join("")}<th class="z">§ 6b-Rücklage</th><th>Status</th><th></th></tr></thead>`;
    const zeilen = zustand.eingabe.verkaeufe.map((roh, j) => {
      const v = R.verkaeufe[j];
      const zellen = VK_FELDER.map((k, n) => {
        let td = eingabeFeld(`vk:${j}:${k}`, felder[k].format, roh[k], v.felder[k], v.herkunft[k] || "", felder[k]);
        if (k === "objekt_id") td = td.replace("<input ", `<input list="liste-objekte" `);
        return n === 0 ? td.replace('<td class="', '<td class="fest ') : td;
      }).join("");
      const rl = R.ruecklagen[j];
      return `<tr>${zellen}${VK_ERGEBNIS.map((k) => zelle(`vk:${j}:${k}`)).join("")}${rl ? zelle(`rl:${j}:betrag`) : "<td></td>"}<td class="status ${statusKlasse(v.status)}">${esc(v.status)}</td><td><button class="knopf klein gefahr" data-aktion="vk-weg" data-nr="${j}">✕</button></td></tr>`;
    }).join("");
    return seitenkopf("Verkäufe", "Verkauf zum Jahresende: Miete und AfA laufen im Verkaufsjahr noch. Leerer Preis = Verkehrswert fortgeschrieben (Annahme). Weitere Kostenstellen werden als Paket mitverkauft.",
      `<button class="knopf haupt" data-aktion="vk-neu">+ Verkauf</button>`)
      + FARBLEGENDE
      + `<datalist id="liste-objekte">${ids.map((id) => `<option value="${esc(id)}">`).join("")}</datalist>`
      + (zeilen ? `<div class="tabelle-rahmen"><table>${kopf}<tbody>${zeilen}</tbody></table></div>` : `<div class="leer-hinweis">Keine Verkäufe geplant.</div>`);
  }

  const NE_GRUPPEN = [
    ["Kauf", ["neu_id", "aktiv", "name", "kaufjahr", "kaufpreis", "anteil_gub", "nebenkosten", "ak_gub_eingabe", "ak_geb_eingabe"]],
    ["AfA", ["afa_satz", "afa_methode", "qng", "qng_grenze"]],
    ["Miete und Kosten", ["mietrendite", "erhaltungsquote", "flaeche", "miete_qm", "mietsteig", "miet_warte", "miet_bis", "kst"]],
    ["§ 6b-Rücklagen", ["quelle", "quelle2", "quelle3"]],
    ["Finanzierung", ["erloes_fin", "fin_art", "kredit_betrag", "kredit_quote", "kredit_zins", "tilgungsart", "tilgung", "laufzeit", "tilgungsfrei"]],
    ["Zweitkredit", ["kredit2_anteil", "kredit2_betrag", "kredit2_zins", "kredit2_tilgungsart", "kredit2_tilgung", "kredit2_laufzeit", "kredit2_tilgungsfrei", "kredit2_uebernahme"]],
  ];
  function ansichtNeuobjekte() {
    const R = zustand.ergebnis;
    const felder = Object.fromEntries(M.felder.neuobjekte.map((f) => [f.key, f]));
    const rlIds = R.ruecklagen.filter(Boolean).map((r) => r.id);
    const karten = R.neuobjekte.map((x, r) => {
      const roh = zustand.eingabe.neuobjekte[r];
      if (!x.berechnet && !roh) return "";
      const erfasst = !!roh;
      const b = x.berechnet || {};
      const titel = `${esc(x.felder.neu_id || "(ohne NeuID)")}${x.felder.name ? " · " + esc(x.felder.name) : ""}`;
      const kenn = x.berechnet ? [
        ["Kaufpreis + Nebenkosten", `ne:${r}:ak_gesamt`], ["übertragen § 6b", `ne:${r}:ue_gesamt`], ["AfA-Basis Gebäude", `ne:${r}:afa_basis`],
        ["Bemessung § 7b", `ne:${r}:qng_basis`], ["Kredit", `ne:${r}:kredit`], ["Eigenmittel", `ne:${r}:eigen`],
      ].map(([l, p]) => `<div>${esc(l)}<b data-pfad="${p}">${esc(format(WEG.wertVon(p, R), "euro"))}</b></div>`).join("") : "";
      const abschnitte = NE_GRUPPEN.map(([gtitel, keys]) => {
        const inhalt = keys.map((k) => {
          const f = felder[k];
          const rohWert = erfasst ? roh[k] : undefined;
          const wirksam = x.felder[k];
          const herkunft = x.herkunft[k] || "";
          if (!erfasst) return `<div class="feld"><label>${esc(f.ueberschrift)}</label><div class="wert">${esc(format(wirksam, f.format))}</div></div>`;
          let html = formFeld(`ne:${r}:${k}`, f, rohWert, wirksam, herkunft);
          if (k.startsWith("quelle")) html = html.replace("<input ", `<input list="liste-ruecklagen" `);
          if (k === "kst") html = html.replace("<input ", `<input list="liste-neukauf" `);
          return html;
        }).join("");
        let zusatz = "";
        if (gtitel === "§ 6b-Rücklagen" && x.berechnet) {
          zusatz = `<div class="tabelle-rahmen" style="margin-top:10px"><table><thead><tr><th class="fest">Quelle</th><th class="z">Gebäude verfügbar</th><th class="z">G+B verfügbar</th><th class="z">ü1 Geb. → Geb.</th><th class="z">ü2 G+B → G+B</th><th class="z">ü3 G+B → Geb.</th><th class="z">Erlös eingesetzt</th></tr></thead><tbody>${[0, 1, 2].filter((k) => !leer(x.felder[RW.konstanten.QUELLEN[k]])).map((k) => `<tr><td class="fest">${esc(x.felder[RW.konstanten.QUELLEN[k]])}</td>${["rl_geb", "rl_gub", "ue1", "ue2", "ue3", "erloes"].map((s) => zelle(`neq:${r}:${k}:${s}`)).join("")}</tr>`).join("") || '<tr><td class="fest leise" colspan="7">keine Quelle: Kauf ohne § 6b-Übertragung</td></tr>'}</tbody></table></div>`;
        }
        const offen = gtitel !== "Zweitkredit" || keys.some((k) => !leer(x.felder[k]));
        return `<section><details${offen ? " open" : ""}><summary><h4 style="display:inline">${esc(gtitel)}</h4></summary><div class="formular" style="margin-top:8px">${inhalt}</div>${zusatz}</details></section>`;
      }).join("");
      return `<div class="vorgang">
        <div class="vorgang-kopf"><div><div class="titel">${titel}</div><div class="leise">Kauf Ende ${esc(x.felder.kaufjahr)}${erfasst ? "" : " · aus „reinvestieren = ja“ im Blatt Verkäufe (Annahmen)"}</div></div>
          <div style="display:flex;gap:10px;align-items:center"><span class="status ${statusKlasse(x.status)}">${esc(x.status)}</span>${erfasst ? `<button class="knopf klein" data-aktion="ne-kopie" data-nr="${r}">Kopieren</button><button class="knopf klein gefahr" data-aktion="ne-weg" data-nr="${r}">Entfernen</button>` : `<button class="knopf klein" data-aktion="ne-uebernehmen" data-nr="${r}">Als eigenes Neuobjekt bearbeiten</button>`}</div></div>
        ${kenn ? `<div class="vorgang-kennzahlen">${kenn}</div>` : ""}
        <div class="abschnitte">${abschnitte}</div></div>`;
    }).join("");
    const nk = (zustand.eingabe.neukauf || []).map((n) => n.kst);
    return seitenkopf("Neuobjekte", "Kauf zum Jahresende: Übertragung der Rücklage und Kreditauszahlung im Kaufjahr, Miete, AfA, Zins und Tilgung ab dem Folgejahr. Bis zu drei Quell-Rücklagen in fester Reihenfolge: ü1 Gebäudegewinn auf Gebäude, ü2 G+B-Gewinn auf G+B, ü3 Rest auf Gebäude.",
      `<button class="knopf haupt" data-aktion="ne-neu">+ Neuobjekt</button>`)
      + FARBLEGENDE
      + `<datalist id="liste-ruecklagen">${rlIds.map((id) => `<option value="${esc(id)}">`).join("")}</datalist><datalist id="liste-neukauf">${nk.map((id) => `<option value="${esc(id)}">`).join("")}</datalist>`
      + (karten || `<div class="leer-hinweis">Keine Neuobjekte. „+ Neuobjekt“ legt einen Kauf an; Pflicht sind NeuID, Kaufjahr, Kaufpreis, AfA-Satz und Anteil G+B.</div>`);
  }

  function ansichtDaten() {
    const R = zustand.ergebnis, E = zustand.eingabe;
    const anl = R.anlagen.filter(Boolean);
    const anlTab = anl.length ? `<div class="tabelle-rahmen"><table><thead><tr><th class="fest">Inventar-Nr.</th><th>Bezeichnung</th><th>ObjektID</th><th>Art</th><th>Methode</th><th class="z">AHK</th><th class="z">Buchwert Stand</th><th class="z">AfA-Satz</th><th class="z">AfA p. a.</th><th class="z">Buchwert Basisjahr</th><th>Status</th></tr></thead><tbody>${R.anlagen.map((a, j) => a ? `<tr><td class="fest">${esc(a.felder.nr)}</td><td>${esc(a.felder.bezeichnung)}</td><td>${esc(a.felder.objekt_id)}</td><td>${esc(a.felder.art)}</td><td>${esc(a.felder.methode)}</td><td class="z">${esc(format(a.felder.ahk, "euro"))}</td><td class="z">${esc(format(a.felder.bw_stand, "euro"))}</td><td class="z">${esc(format(a.felder.satz, "prozent"))}</td><td class="z">${esc(format(a.afa, "euro"))}</td><td class="z">${esc(format(a.bw_basis, "euro"))}</td><td class="status ${statusKlasse(a.status)}">${esc(a.status)}</td></tr>` : "").join("")}</tbody></table></div>` : `<div class="leer-hinweis">Kein Anlagenverzeichnis.</div>`;
    const plan = Object.entries(E.afa_plan || {});
    const planTab = plan.length ? `<div class="tabelle-rahmen"><table>${jahresKopf(R.jahre, false, "ObjektID")}<tbody>${plan.map(([id, w]) => `<tr><td class="fest">${esc(id)}</td>${R.jahre.map((t) => `<td class="z">${esc(format(w[t], "euro"))}</td>`).join("")}</tr>`).join("")}</tbody></table></div>` : `<div class="leer-hinweis">Kein AfA-Plan.</div>`;
    const nk = E.neukauf || [];
    const nkJahre = [R.parameter.Basisjahr, ...R.jahre];
    const nkTab = nk.length ? `<div class="tabelle-rahmen"><table>${jahresKopf(nkJahre, false, "Kostenstelle · Position")}<tbody>${nk.map((n) => M.neukauf_positionen.map((p) => {
      const fort = (R.neukauf[String(n.kst).toLowerCase()] || {})[p.key] || [];
      return `<tr><td class="fest">${esc(n.kst)} · ${esc(p.bezeichnung)}</td>${nkJahre.map((t, i) => `<td class="z${istZahl((n.werte[p.key] || {})[t]) ? "" : " leise"}">${esc(format(fort[i], "euro"))}</td>`).join("")}</tr>`;
    }).join("")).join("")}</tbody></table></div>` : `<div class="leer-hinweis">Keine Neukauf-Kostenstellen.</div>`;
    return seitenkopf("Anlagen und Planwerte", "Diese Daten kommen aus der Excel-Mappe (Anlagenverzeichnis, AfA-Plan aus BWA 1240, Neukauf-Kostenstellen) und werden hier nur angezeigt. Ändern in der Mappe, dann die Seite neu erzeugen: python -m prognosemodell.browser --mappe Prognose.xlsx.")
      + `<h3>Anlagenverzeichnis</h3>${anlTab}<h3>AfA-Plan</h3>${planTab}<h3>Neukauf-Kostenstellen (fortgeschrieben, grau = Annahme bzw. Fortschreibung)</h3>${nkTab}`;
  }

  // ---------------------------------------------------------------- Ansichten: Rechenweg
  const PRG_BESTAND = [
    { abschnitt: "Laufendes Ergebnis" }, "aktiv", "miete", "einnahmen", "erhaltung", "ausgaben", "afa", "ergebnis",
    { abschnitt: "Buchwerte und Wert am Jahresende" }, "buchwert", "buchwert_gub", "verkehrswert", "bestand",
    { abschnitt: "Darlehen" }, "kredit_zins", "tilgung", "abloesung", "restschuld",
    { abschnitt: "Bei Halten (ohne Verkauf)" }, "miete_halten", "einnahmen_halten", "erhaltung_halten", "afa_halten", "buchwert_halten", "zins_halten", "tilgung_halten", "restschuld_halten", "ergebnis_halten", "wirkung",
  ];
  const PRG_NEU = [
    { abschnitt: "Laufendes Ergebnis" }, "aktiv", "miete", "einnahmen", "erhaltung", "ausgaben", "afa", "afa_qng", "ergebnis",
    { abschnitt: "Buchwerte und Wert am Jahresende" }, "buchwert", "buchwert_gub", "verkehrswert", "bestand",
    { abschnitt: "Darlehen" }, "kredit_zins", "tilgung", "restschuld",
    { abschnitt: "Ohne § 6b-Übertragung (Szenarien B und C)" }, "afa_ohne6b", "afa_qng_ohne6b", "buchwert_ohne6b", "buchwert_gub_ohne6b",
  ];
  function ansichtPrognose() {
    const R = zustand.ergebnis;
    const optionen = [];
    R.prognose.bestand.forEach((s, j) => { if (s) optionen.push([`b:${j}`, `${s.id[0]}${R.objekte[j].felder.name ? " · " + R.objekte[j].felder.name : ""}`]); });
    R.prognose.neu.forEach((s, r) => { if (s) optionen.push([`n:${r}`, `${s.id[0]} (Neuobjekt)`]); });
    if (!optionen.length) return seitenkopf("Prognose je Objekt", "") + `<div class="leer-hinweis">Keine Objekte.</div>`;
    if (!optionen.some(([k]) => k === zustand.prgObjekt)) zustand.prgObjekt = optionen[0][0];
    const [art, nr] = zustand.prgObjekt.split(":");
    const sp = Object.fromEntries(M.spalten.prognose.map((s) => [s.key, s]));
    const zeilen = (art === "b" ? PRG_BESTAND : PRG_NEU).map((k) => typeof k === "object" ? k : ({
      label: sp[k].ueberschrift, pfad: (i) => `prg:${art}:${nr}:${k}:${i}`,
      summe: !["aktiv", "bestand", "buchwert", "buchwert_gub", "verkehrswert", "restschuld", "buchwert_halten", "restschuld_halten", "buchwert_ohne6b", "buchwert_gub_ohne6b"].includes(k),
      klasse: k === "ergebnis" ? "summe" : k === "afa_qng" || k === "abloesung" ? "unter" : "",
    }));
    const auswahl = `<select data-aktion="prg-objekt" class="knopf">${optionen.map(([k, t]) => `<option value="${k}"${k === zustand.prgObjekt ? " selected" : ""}>${esc(t)}</option>`).join("")}</select>`;
    let info = "";
    if (art === "b") {
      const o = R.objekte[Number(nr)];
      const vk = R.verkaeufe.find((v) => v.aktiv && String(v.umfang).toLowerCase().includes(";" + String(o.felder.objekt_id).toLowerCase() + ";") && !leer(v.felder.jahr));
      info = `${vk ? `Verkauf Ende ${esc(vk.felder.jahr)}` : "kein Verkauf geplant"} · Restbuchwert Basisjahr ${esc(format(o.felder.restbuchwert, "euro"))} · AfA-Satz ${esc(format(o.felder.afa_satz, "prozent"))} · Status ${esc(o.status)}`;
    } else {
      const x = R.neuobjekte[Number(nr)];
      info = `Kauf Ende ${esc(x.felder.kaufjahr)} · AfA-Basis ${esc(format(x.berechnet.afa_basis, "euro"))} · ${esc(x.felder.afa_methode || "linear")} ${esc(format(x.felder.afa_satz, "prozent"))}${String(x.felder.qng).toLowerCase() === "ja" ? " · Sonder-AfA § 7b" : ""} · Status ${esc(x.status)}`;
    }
    return seitenkopf("Prognose je Objekt", "Je Jahr Miete, Kosten, AfA, Buchwerte und Darlehen. Was bei Halten stünde, steht darunter; der Plan nimmt es bis zum Verkaufsjahr.", auswahl)
      + `<div class="hinweisbox">${info}</div>` + jahresTabelle(zeilen, true);
  }

  function ansichtRuecklagen() {
    const R = zustand.ergebnis;
    const sp = M.spalten.ruecklage.filter((s) => !["id", "hinweis"].includes(s.key));
    const zeilen = R.ruecklagen.map((r, j) => r ? `<tr><td class="fest">${esc(r.id)}</td>${sp.map((s) => s.format === "text" ? `<td>${esc(r[s.key])}</td>` : zelle(`rl:${j}:${s.key}`)).join("")}<td class="leise">${esc(r.hinweis)}</td></tr>` : "").join("");
    const tab = zeilen ? `<div class="tabelle-rahmen"><table><thead><tr><th class="fest">RücklageID</th>${sp.map((s) => `<th class="${s.format === "text" ? "" : "z"}">${esc(s.ueberschrift)}</th>`).join("")}<th>Hinweis</th></tr></thead><tbody>${zeilen}</tbody></table></div>` : `<div class="leer-hinweis">Keine § 6b-Rücklage: kein Verkauf mit „§ 6b nutzen = ja“ und Status OK.</div>`;
    const spiegel = M.spalten.spiegel.filter((s) => s.key !== "jahr").map((s) => ({ label: s.ueberschrift, pfad: (i) => `rls:${s.key}:${i}`, summe: s.key !== "bestand", klasse: s.key === "steuerwirksam" || s.key === "bestand" ? "summe" : "" }));
    return seitenkopf("§ 6b-Rücklagen", "Je Verkauf mit § 6b eine Rücklage aus den positiven Teilgewinnen, getrennt nach Gebäude und G+B. Was bis zum Fristjahr nicht auf Neuobjekte übertragen ist, wird dort mit Gewinnzuschlag aufgelöst.")
      + tab + `<h3>Spiegel je Jahr</h3>` + jahresTabelle(spiegel, true);
  }

  function ansichtDarlehen() {
    const R = zustand.ergebnis;
    const bloecke = [];
    R.darlehen.forEach((d, r) => {
      if (!d) return;
      ["haupt", "zweit"].forEach((teil) => {
        const k = d[teil];
        if (!k) return;
        const info = `${esc(k.id)} · ${esc(k.teil)} · ${esc(format(k.betrag, "euro"))} zu ${esc(format(k.zinssatz, "prozent"))} · ${esc(k.art)}${k.tilgungssatz ? ", Tilgung " + esc(format(k.tilgungssatz, "prozent")) : ""}${istZahl(k.laufzeit) ? ", Laufzeit " + k.laufzeit + " Jahre" : ""}${k.tilgungsfrei ? ", " + k.tilgungsfrei + " Jahre tilgungsfrei" : ""} · Rate ${esc(format(k.rate, "euro"))} · getilgt ${esc(k.getilgt)}`;
        bloecke.push(`<h3>${info}</h3>` + jahresTabelle([
          { label: "Zinsen", pfad: (i) => `dl:${r}:${teil}:zins:${i}` },
          { label: "Tilgung", pfad: (i) => `dl:${r}:${teil}:tilgung:${i}` },
          { label: "Restschuld Ende", pfad: (i) => `dl:${r}:${teil}:restschuld:${i}`, summe: false, klasse: "summe" },
        ], true));
      });
    });
    const bestand = [];
    R.prognose.bestand.forEach((s, j) => {
      if (!s || !(s.restschuld.some((w) => w > 0) || s.kredit_zins.some((w) => w > 0))) return;
      bestand.push(`<h3>${esc(s.id[0])} (Bestand)</h3>` + jahresTabelle([
        { label: "Zinsen", pfad: (i) => `prg:b:${j}:kredit_zins:${i}` },
        { label: "Tilgung (mit Ablösung)", pfad: (i) => `prg:b:${j}:tilgung:${i}` },
        { label: "Restschuld Ende", pfad: (i) => `prg:b:${j}:restschuld:${i}`, summe: false, klasse: "summe" },
      ], true));
    });
    return seitenkopf("Darlehen", "Kredite der Neuobjekte: Auszahlung zum Ende des Kaufjahrs, Zins und Tilgung ab dem Folgejahr auf die Restschuld am Vorjahresende. Darlehen der Bestandsobjekte: Tilgungsplan aus Restschuld, Zinssatz und Rate, beim Verkauf aus dem Erlös abgelöst.")
      + (bloecke.join("") || `<div class="leer-hinweis">Keine Kredite bei Neuobjekten.</div>`)
      + (bestand.length ? bestand.join("") : "");
  }

  // ---------------------------------------------------------------- Prüfung
  function pruefungGesamt(R) {
    const liste = M.pruefungen.filter((p) => !OHNE_BROWSER.includes(p.key));
    const fehler = liste.filter((p) => p.art === "Fehler" && R.pruefung[p.key] > 0).length;
    const warn = liste.filter((p) => p.art === "Warnung" && R.pruefung[p.key] > 0).length;
    if (fehler) return { text: `${fehler} Fehler`, klasse: "fehler" };
    if (warn) return { text: `${warn} Warnung(en)`, klasse: "warnung" };
    return { text: "OK", klasse: "ok" };
  }
  function ansichtPruefung() {
    const R = zustand.ergebnis;
    const rang = { Fehler: 0, Warnung: 1, Hinweis: 2 };
    const liste = M.pruefungen.filter((p) => !OHNE_BROWSER.includes(p.key))
      .map((p) => ({ ...p, n: R.pruefung[p.key] || 0 }))
      .sort((a, b) => (b.n > 0) - (a.n > 0) || rang[a.art] - rang[b.art]);
    const status = [];
    R.objekte.forEach((o) => { if (o && o.status !== "OK") status.push(["objekte", "Objekt " + o.felder.objekt_id, o.status]); });
    R.verkaeufe.forEach((v) => { if (v.status && v.status !== "OK") status.push(["verkaeufe", "Verkauf " + v.felder.objekt_id, v.status]); });
    R.neuobjekte.forEach((x) => { if (x.status && x.status !== "OK") status.push(["neuobjekte", "Neuobjekt " + x.felder.neu_id, x.status]); });
    const g = pruefungGesamt(R);
    return seitenkopf("Prüfung", `Plausibilitätsprüfungen wie im Blatt Prüfung der Mappe. Ergebnis: <b>${esc(g.text)}</b>.`)
      + (status.length ? `<h3>Zeilen mit Status ungleich OK</h3><div class="pruefliste">${status.map(([a, w, s]) => `<div class="pruefung"><div class="art Fehler">Status</div><div class="zahl"></div><div><a href="#" data-ansicht="${a}">${esc(w)}</a><div class="wo">${esc(s)}</div></div></div>`).join("")}</div><h3>Alle Prüfungen</h3>` : "")
      + `<div class="pruefliste">${liste.map((p) => `<div class="pruefung${p.n ? "" : " ok"}"><div class="art ${p.n ? p.art : "OK"}">${p.n ? p.art : "OK"}</div><div class="zahl">${p.n}</div><div>${esc(p.bezeichnung)}<div class="wo">${esc(p.wo)}</div></div></div>`).join("")}</div>`;
  }

  // ---------------------------------------------------------------- Rechenweg-Leiste
  function zeigeWeg(pfad, neu = true) {
    if (neu) zustand.weg = [pfad];
    else if (zustand.weg[zustand.weg.length - 1] !== pfad) zustand.weg.push(pfad);
    zustand.gewaehlt = pfad;
    zeichneWeg();
    document.querySelectorAll("td.gewaehlt").forEach((td) => td.classList.remove("gewaehlt"));
    document.querySelectorAll(`[data-pfad="${CSS.escape(pfad)}"]`).forEach((td) => td.classList.add("gewaehlt"));
  }
  function zeichneWeg() {
    const el = $("#rechenweg");
    if (!zustand.weg.length || !zustand.ergebnis) { el.hidden = true; return; }
    const pfad = zustand.weg[zustand.weg.length - 1];
    const R = zustand.ergebnis;
    const e = WEG.erklaere(pfad, R, M);
    const art = artVon(pfad);
    const verlauf = zustand.weg.length > 1 ? `<div class="verlauf">${zustand.weg.slice(0, -1).map((p, n) => `<button data-verlauf="${n}" title="${esc(WEG.bezeichnung(p, M))}">${esc(WEG.bezeichnung(p, M))}</button>`).join("<span class=leise>›</span>")}</div>` : "";
    const teile = e.teile.map((t) => {
      const fmt = t.pfad ? artVon(t.pfad) : t.format || "euro";
      const wer = t.wer && t.wer !== e.wer ? `<small>${esc(t.wer)}</small>` : "";
      const jahr = t.pfad && /:(\d+)$/.test(t.pfad) && /^(prg|liq|aus|rls|dl)/.test(t.pfad) ? R.jahre[Number(t.pfad.split(":").pop())] : null;
      const jahrText = jahr && jahr !== e.jahr && String(jahr) !== String(t.label) ? `<small>${jahr}</small>` : "";
      return `<div class="teil${t.pfad ? " link" : ""}"${t.pfad ? ` data-weg="${esc(t.pfad)}"` : ""}><span class="t-label"><span class="t-name">${esc(t.label)}</span>${wer}${jahrText}</span><span class="t-wert">${esc(format(t.wert, fmt)) || "–"}</span></div>`;
    }).join("");
    el.innerHTML = `${verlauf}<div class="weg-kopf"><div><div class="weg-titel">${esc(e.titel)}</div><div class="weg-wer">${esc([e.wer, e.jahr].filter(Boolean).join(" · "))}</div></div><button class="schliessen" data-aktion="weg-zu" title="schließen">×</button></div>
      <div class="weg-wert">${esc(format(e.wert, art)) || "–"}</div>
      ${e.formel ? `<div class="formel"><b>Rechnung</b>${esc(e.formel)}</div>` : `<div class="formel"><b>Herkunft</b>${/^(obj|vk|ne|par):/.test(pfad) ? "Eingabe" : "berechnet"}</div>`}
      ${teile ? `<div class="teile">${teile}</div>` : ""}
      <div class="excel">${e.excel ? `In der Excel-Mappe: <code>${esc(e.excel)}</code>` : ""}${zielAnsicht(pfad) && zielAnsicht(pfad) !== zustand.ansicht ? ` · <a href="#" data-zeige="${esc(pfad)}">im Blatt zeigen</a>` : ""}</div>`;
    el.hidden = false;
  }
  function zielAnsicht(pfad) {
    const t = pfad.split(":");
    return { par: "parameter", obj: "objekte", prg: "prognose", vk: "verkaeufe", rl: "ruecklagen", rls: "ruecklagen", ne: "neuobjekte",
      neq: "neuobjekte", dl: "darlehen", liq: "liquiditaet", aus: "vermoegen", vg: "vergleich" }[t[0]];
  }
  function zeigeImBlatt(pfad) {
    const t = pfad.split(":");
    const ziel = zielAnsicht(pfad);
    if (!ziel) return;
    if (t[0] === "prg") zustand.prgObjekt = `${t[1]}:${t[2]}`;
    if (t[0] === "liq" || t[0] === "aus") zustand.sz = t[1];
    zustand.ansicht = ziel;
    zeichnen();
    const td = document.querySelector(`[data-pfad="${CSS.escape(pfad)}"]`);
    if (td) td.scrollIntoView({ block: "center", inline: "center" });
  }

  // ---------------------------------------------------------------- Zeichnen
  function zeichnen() {
    zeichneNavi();
    const R = zustand.ergebnis;
    const E = zustand.eingabe;
    const info = [];
    if (E.quelle) info.push("Daten aus " + E.quelle);
    if (zustand.quelle === "entwurf") info.push("Entwurf aus diesem Browser");
    info.push(`erzeugt ${window.PROGNOSE_ERZEUGT || ""}`);
    if (zustand.dauer !== undefined) info.push(`gerechnet in ${F0.format(Math.max(zustand.dauer, 1))} ms`);
    $("#kopf-info").textContent = info.join(" · ");
    const chip = $("#btn-pruefung");
    if (R) {
      const g = pruefungGesamt(R);
      chip.textContent = "Prüfung: " + g.text;
      chip.className = "chip " + g.klasse;
    }
    const inhalt = $("#inhalt");
    const scroll = { top: inhalt.scrollTop };
    const rahmen = inhalt.querySelector(".tabelle-rahmen");
    if (rahmen) { scroll.rt = rahmen.scrollTop; scroll.rl = rahmen.scrollLeft; }
    if (zustand.fehler || !R) {
      inhalt.innerHTML = `<div class="hinweisbox">Die Rechnung ist abgebrochen: ${esc(zustand.fehler && zustand.fehler.message)}. Eingaben prüfen oder „Daten der Datei wiederherstellen“.</div>`;
      return;
    }
    const ansicht = {
      ueberblick: ansichtUeberblick, vergleich: ansichtVergleich, liquiditaet: ansichtLiquiditaet, vermoegen: ansichtVermoegen,
      parameter: ansichtParameter, objekte: ansichtObjekte, verkaeufe: ansichtVerkaeufe, neuobjekte: ansichtNeuobjekte,
      daten: ansichtDaten, prognose: ansichtPrognose, ruecklagen: ansichtRuecklagen, darlehen: ansichtDarlehen, pruefung: ansichtPruefung,
    }[zustand.ansicht] || ansichtUeberblick;
    const aktiv = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.ziel : null;
    inhalt.innerHTML = ansicht();
    if (zustand.ansichtAlt === zustand.ansicht) {
      inhalt.scrollTop = scroll.top;
      const r2 = inhalt.querySelector(".tabelle-rahmen");
      if (r2 && scroll.rt !== undefined) { r2.scrollTop = scroll.rt; r2.scrollLeft = scroll.rl; }
    } else inhalt.scrollTop = 0;
    zustand.ansichtAlt = zustand.ansicht;
    if (aktiv) {
      const el = inhalt.querySelector(`[data-ziel="${CSS.escape(aktiv)}"]`);
      if (el && zustand.naechstesFeld) {
        const felder = [...inhalt.querySelectorAll("[data-ziel]")];
        const n = felder.findIndex((f) => f.dataset.ziel === aktiv);
        const z = felder[n + 1];
        if (z) z.focus();
      }
    }
    zustand.naechstesFeld = false;
    zeichneWeg();
  }

  // ---------------------------------------------------------------- Eingaben ändern
  function aendern(ziel, text, art) {
    const t = ziel.split(":");
    const wert = lese(text, art);
    const E = zustand.eingabe;
    if (t[0] === "par") {
      const vorgabe = (M.parameter.find((p) => p.name === t[1]) || {}).wert;
      E.parameter[t[1]] = wert === null ? (vorgabe === undefined ? null : vorgabe) : wert;
    } else {
      const liste = { obj: E.objekte, vk: E.verkaeufe, ne: E.neuobjekte }[t[0]];
      const zeile = liste[Number(t[1])];
      zeile[t[2]] = wert;
      if (zeile.herkunft) delete zeile.herkunft[t[2]];
    }
    neuRechnen();
  }
  function leeresObjekt(liste, praefix, schluessel, felder) {
    let n = liste.length + 1;
    const ids = liste.map((x) => String(x[schluessel]));
    while (ids.includes(`${praefix}${String(n).padStart(3, "0")}`)) n++;
    const o = Object.fromEntries(felder.map((f) => [f.key, null]));
    o[schluessel] = `${praefix}${String(n).padStart(3, "0")}`;
    return o;
  }
  function aktion(name, nr) {
    const E = zustand.eingabe;
    switch (name) {
      case "obj-neu": E.objekte.push(leeresObjekt(E.objekte, "OBJ-", "objekt_id", M.felder.objekte)); break;
      case "obj-weg": if (confirmFrage(`Objekt ${E.objekte[nr].objekt_id || ""} entfernen?`)) E.objekte.splice(nr, 1); else return; break;
      case "vk-neu": {
        const v = Object.fromEntries(M.felder.verkaeufe.filter((f) => f.key !== "bezeichnung").map((f) => [f.key, null]));
        const erstes = E.objekte.find((o) => !leer(o.objekt_id));
        v.objekt_id = erstes ? erstes.objekt_id : "";
        v.jahr = zustand.ergebnis.jahre[2];
        E.verkaeufe.push(v);
        break;
      }
      case "vk-weg": E.verkaeufe.splice(nr, 1); break;
      case "ne-neu": {
        const n = Object.fromEntries(M.felder.neuobjekte.map((f) => [f.key, null]));
        Object.assign(n, { neu_id: leeresObjekt(E.neuobjekte.filter(Boolean), "NEU-", "neu_id", []).neu_id,
          kaufjahr: zustand.ergebnis.jahre[3], kaufpreis: 1000000, anteil_gub: 0.2, afa_satz: 0.03, mietrendite: 0.045, erhaltungsquote: 0.005 });
        // Annahmezeilen aus Verkäufen nicht überschreiben: hinten anhängen
        const frei = Math.max(E.neuobjekte.length, zustand.ergebnis.neuobjekte.length);
        while (E.neuobjekte.length < frei) E.neuobjekte.push(null);
        E.neuobjekte.push(n);
        break;
      }
      case "ne-weg": E.neuobjekte[nr] = null; while (E.neuobjekte.length && E.neuobjekte[E.neuobjekte.length - 1] === null) E.neuobjekte.pop(); break;
      case "ne-kopie": { const k = kopie(E.neuobjekte[nr]); k.neu_id = String(k.neu_id) + "-2"; E.neuobjekte.push(k); break; }
      case "ne-uebernehmen": {
        const x = zustand.ergebnis.neuobjekte[nr];
        const n = Object.fromEntries(M.felder.neuobjekte.map((f) => [f.key, leer(x.felder[f.key]) ? null : x.felder[f.key]]));
        n.ak_gub_eingabe = null; n.ak_geb_eingabe = null;
        while (E.neuobjekte.length < nr) E.neuobjekte.push(null);
        E.neuobjekte[nr] = n;
        break;
      }
      case "weg-zu": zustand.weg = []; zustand.gewaehlt = null; zeichneWeg(); document.querySelectorAll("td.gewaehlt").forEach((td) => td.classList.remove("gewaehlt")); return;
      default: return;
    }
    neuRechnen();
  }
  function confirmFrage(text) { return window.confirm(text); }

  // ---------------------------------------------------------------- Dateien
  function herunterladen() {
    const daten = { format: "prognosemodell-eingabe", version: 1, gespeichert: new Date().toISOString(), eingabe: zustand.eingabe };
    const blob = new Blob([JSON.stringify(daten, null, 1)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `Prognose_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function oeffnen(datei) {
    const leser = new FileReader();
    leser.onload = () => {
      try {
        const roh = JSON.parse(leser.result);
        const e = roh && roh.eingabe ? roh.eingabe : roh;
        if (!e || !Array.isArray(e.objekte)) throw new Error("keine Eingabedatei des Prognosemodells");
        e.parameter = Object.assign(Object.fromEntries(M.parameter.filter((p) => !p.berechnet).map((p) => [p.name, p.wert])), e.parameter || {});
        for (const k of ["verkaeufe", "neuobjekte", "anlagen", "neukauf"]) e[k] = e[k] || [];
        e.afa_plan = e.afa_plan || {};
        e.quelle = datei.name;
        setzeEingabe(e, "datei");
      } catch (err) {
        alert("Die Datei lässt sich nicht öffnen: " + err.message);
      }
    };
    leser.readAsText(datei);
  }
  function beispielEingabe() {
    return window.PROGNOSE_BEISPIEL ? kopie(window.PROGNOSE_BEISPIEL) : kopie(START);
  }

  // ---------------------------------------------------------------- Ereignisse
  function verbinden() {
    $("#navi").addEventListener("click", (e) => {
      const b = e.target.closest("[data-ansicht]");
      if (!b) return;
      zustand.ansicht = b.dataset.ansicht;
      zeichnen();
    });
    const inhalt = $("#inhalt");
    inhalt.addEventListener("click", (e) => {
      const nav = e.target.closest("[data-ansicht]");
      if (nav) { e.preventDefault(); zustand.ansicht = nav.dataset.ansicht; zeichnen(); return; }
      const sz = e.target.closest("[data-sz]");
      if (sz) { zustand.sz = sz.dataset.sz; zeichnen(); return; }
      const ak = e.target.closest("button[data-aktion]");
      if (ak) { aktion(ak.dataset.aktion, Number(ak.dataset.nr)); return; }
      const p = e.target.closest("[data-pfad]");
      if (p) { zeigeWeg(p.dataset.pfad); return; }
      const svg = e.target.closest("svg[data-diagramm]");
      if (svg && svg.dataset.index !== undefined) {
        const id = svg.dataset.diagramm;
        const i = Number(svg.dataset.index);
        zeigeWeg(id === "vermoegen" ? `aus:A:vermoegen_netto:${i}` : `liq:A:kum:${i}`);
      }
    });
    inhalt.addEventListener("change", (e) => {
      const el = e.target;
      if (el.dataset.ziel) { aendern(el.dataset.ziel, el.value, el.dataset.art); return; }
      if (el.dataset.aktion === "prg-objekt") { zustand.prgObjekt = el.value; zeichnen(); }
    });
    inhalt.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && e.target.dataset && e.target.dataset.ziel && e.target.tagName === "INPUT") {
        zustand.naechstesFeld = true;
        e.target.blur();
      }
    });
    inhalt.addEventListener("mousemove", diagrammHover);
    inhalt.addEventListener("mouseout", diagrammWeg);
    $("#rechenweg").addEventListener("click", (e) => {
      const t = e.target.closest("[data-weg]");
      if (t) { zeigeWeg(t.dataset.weg, false); return; }
      const v = e.target.closest("[data-verlauf]");
      if (v) { zustand.weg = zustand.weg.slice(0, Number(v.dataset.verlauf) + 1); zeigeWeg(zustand.weg[zustand.weg.length - 1], false); return; }
      const z = e.target.closest("[data-zeige]");
      if (z) { e.preventDefault(); zeigeImBlatt(z.dataset.zeige); return; }
      const a = e.target.closest("[data-aktion]");
      if (a) aktion(a.dataset.aktion);
    });
    $("#btn-pruefung").addEventListener("click", () => { zustand.ansicht = "pruefung"; zeichnen(); });
    $("#btn-speichern").addEventListener("click", herunterladen);
    $("#btn-oeffnen").addEventListener("click", () => $("#datei").click());
    $("#datei").addEventListener("change", (e) => { if (e.target.files[0]) oeffnen(e.target.files[0]); e.target.value = ""; });
    const menue = (fn) => () => { document.querySelector(".menue").open = false; fn(); };
    $("#btn-beispiel").addEventListener("click", menue(() => setzeEingabe(beispielEingabe(), "beispiel")));
    $("#btn-start").addEventListener("click", menue(() => setzeEingabe(kopie(START), "datei")));
    $("#btn-leer").addEventListener("click", menue(() => {
      const e = kopie(START);
      Object.assign(e, { objekte: [], verkaeufe: [], neuobjekte: [], anlagen: [], afa_plan: {}, neukauf: [], quelle: null });
      setzeEingabe(e, "leer");
    }));
    $("#btn-drucken").addEventListener("click", menue(() => window.print()));
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && zustand.weg.length) aktion("weg-zu");
    });
  }

  // ---------------------------------------------------------------- Start
  function start() {
    const entwurf = laden();
    zustand.eingabe = entwurf || kopie(START);
    zustand.quelle = entwurf ? "entwurf" : "datei";
    verbinden();
    neuRechnen();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
  else start();
})();
