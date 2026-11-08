#!/usr/bin/env node
// Rechnet eine Eingabe (JSON auf stdin) mit dem Rechenwerk und schreibt das Ergebnis als
// JSON auf stdout. Für den Abgleich mit der Excel-Mappe (pruefung/abgleich.py).
"use strict";
const { rechne } = require("./rechenwerk.js");
let roh = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (d) => (roh += d));
process.stdin.on("end", () => {
  const eingaben = JSON.parse(roh);
  const liste = Array.isArray(eingaben) ? eingaben : [eingaben];
  const ergebnisse = liste.map((e) => {
    try {
      return rechne(e);
    } catch (fehler) {
      return { fehler: String(fehler && fehler.stack || fehler) };
    }
  });
  process.stdout.write(JSON.stringify(Array.isArray(eingaben) ? ergebnisse : ergebnisse[0]));
});
