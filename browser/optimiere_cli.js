#!/usr/bin/env node
// Optimiert Eingaben (JSON auf stdin: [{eingabe, einstellungen, meta}]) und schreibt je Fall
// den Lauf, die optimierten Neuobjekte, die tragbare Ausschüttung danach und die Varianten
// Ausgang und Ergebnis als JSON auf stdout. Für pruefung/pruefen_optimierung.py.
"use strict";
const { rechne } = require("./rechenwerk.js");
const O = require("./optimierung.js");
let roh = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (d) => (roh += d));
process.stdin.on("end", () => {
  const ergebnisse = JSON.parse(roh).map(({ eingabe, einstellungen, meta }) => {
    try {
      const vorher = JSON.stringify(eingabe);
      const lauf = O.optimiere(eingabe, einstellungen, meta);
      const R = lauf.ergebnis;
      return {
        lauf: lauf.lauf,
        unveraendert: JSON.stringify(eingabe) === vorher,
        neuobjekte: R.neuobjekte.map((x) => ({ felder: x.felder, berechnet: x.berechnet })),
        ruecklagen: R.ruecklagen,
        kennzahlen: O.kennzahlen(R, meta.fehlerKeys),
        tragbar: O.tragbareAusschuettung(lauf.eingabe).betrag,
        ausschuettung: lauf.eingabe.parameter.Ausschuettung,
        ausgang: O.variante("Optimierung Ausgang", eingabe, rechne(eingabe), meta.pruefungen),
        variante: O.variante("Optimierung Ergebnis", lauf.eingabe, R, meta.pruefungen),
      };
    } catch (fehler) {
      return { fehler: String((fehler && fehler.stack) || fehler) };
    }
  });
  process.stdout.write(JSON.stringify(ergebnisse));
});
