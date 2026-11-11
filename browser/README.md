# Browser-Fassung

Eine HTML-Datei mit Rechenwerk, Oberfläche und Daten. Erzeugen:

```bash
python -m prognosemodell.browser                        # ausgabe/Prognosemodell.html mit Beispiel
python -m prognosemodell.browser --mappe Prognose.xlsx  # mit den Eingaben einer Mappe
```

| Datei | Inhalt |
| --- | --- |
| `rechenwerk.js` | `rechne(eingabe)`: die Formeln der Mappe in JavaScript, ohne Abhängigkeiten |
| `rechenweg.js` | `erklaere(pfad, ergebnis, meta)`: Formel in Worten und Teile je Zahl |
| `oberflaeche.js`, `stil.css`, `seite.html` | Oberfläche; `prognosemodell/browser.py` setzt alles zu einer Datei zusammen |
| `stil_werte.css` | Farben, Schriften, Rundungen: die Werte eines Stils der Knowledge-Base, erzeugt |
| `optimierung.js` | Optimierung und Varianten: `optimiere`, `tragbareAusschuettung`, `kennzahlen`, `variante`; das Gegenstück zu modOptimierung und modVarianten |
| `rechne_cli.js` | rechnet JSON von stdin, für `python -m pruefung.abgleich` |
| `optimiere_cli.js` | optimiert JSON von stdin, für `python -m pruefung.pruefen_optimierung` |

**Eingabe** (JSON, auch das Format von „Speichern“ unter `eingabe`):

- `parameter`: Namen ohne `par_`, z. B. `{"Basisjahr": 2026, "Wertsteig": 0.02}`; berechnete Parameter (Startjahr, SteuersatzEff, …) rechnet das Rechenwerk.
- `objekte`, `verkaeufe`, `neuobjekte`, `anlagen`: Listen mit den Feldschlüsseln aus `modelle.py` (`OBJEKT_FELDER`, …). `null` = leer gelassen, die Annahme greift wie in der Mappe; `""` = bewusst leer. Ein Neuobjekt `null` ist eine Annahmezeile (Reinvestition aus dem Verkauf derselben Zeile).
- `afa_plan`: `{ObjektID: {Jahr: AfA}}`; `neukauf`: `[{kst, name, werte: {miete|einnahmen|erhaltung|ausgaben: {Jahr: Wert}}}]`.

**Datei von „Speichern“:** neben `eingabe` stehen `varianten` (je Variante die Kennzahlen und eine Kopie der Eingaben zum Einspielen) und `optimierung` (`einstellungen`, letzter `lauf`, `tragbar`). Im Browser liegt dasselbe im lokalen Speicher.

**Pfade** der Zahlen im Rechenweg: `par:Name`, `obj:Zeile:Feld`, `prg:b|n:Zeile:Spalte:Jahr`, `vk:Zeile:Spalte`, `rl:Zeile:Spalte`, `rls:Spalte:Jahr`, `ne:Zeile:Spalte`, `neq:Zeile:Quelle:Spalte`, `dl:Zeile:haupt|zweit:Spalte[:Jahr]`, `liq|aus:Szenario:Spalte:Jahr`, `vg:Kennzahl:Szenario`. Zeile und Jahr zählen ab 0.

Wer eine Formel in `prognosemodell/formeln.py` ändert, ändert sie in `rechenwerk.js` mit und prüft mit `python -m pruefung.abgleich`. Wer `modOptimierung.bas` oder die Kennzahlen des Blatts Optimierung ändert, ändert `optimierung.js` mit und prüft mit `python -m pruefung.pruefen_optimierung`.

## Aussehen

Die Seite trägt den Stil „Stil nach awi“ der Knowledge-Base (`design.stil-nach-karriere-bei-awi-tax-gmbh-co-k`). `stil_werte.css` hält dessen Werte, `stil.css` ordnet oben die Rollen der Seite den Rollen des Stils zu (`--akzent: var(--accent)`, Szenario A bis Halten = `--s1` bis `--s4`) und enthält keine eigenen Farbwerte. Nach einer Änderung im Designer, oder für einen anderen Stil:

```bash
python -m kb style design.stil-nach-karriere-bei-awi-tax-gmbh-co-k --css --out stil.css   # im Repository Knowledge-Base
python -m prognosemodell.browser --stil ../Knowledge-Base/stil.css                         # schreibt browser/stil_werte.css
```

Vom Stil und seinen Hausregeln weicht die Seite an zwei Stellen ab:

- **Schriftgrößen und Abstände** bleiben die eigenen, dichten Werte der Seite (14 px Grundschrift) statt der Skala `--t-*` und `--sp-*`: Tabellen über 20 Jahre brauchen die Dichte.
- **Eingabezellen** nutzen `--warn-bg`, weil der Stil keine Rolle dafür hat; es kommt dem Gelb der Mappe am nächsten.

Die Excel-Mappe behält ihre eigenen Farben.
