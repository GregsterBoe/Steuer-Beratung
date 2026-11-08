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
| `rechne_cli.js` | rechnet JSON von stdin, für `python -m pruefung.abgleich` |

**Eingabe** (JSON, auch das Format von „Speichern“ unter `eingabe`):

- `parameter`: Namen ohne `par_`, z. B. `{"Basisjahr": 2026, "Wertsteig": 0.02}`; berechnete Parameter (Startjahr, SteuersatzEff, …) rechnet das Rechenwerk.
- `objekte`, `verkaeufe`, `neuobjekte`, `anlagen`: Listen mit den Feldschlüsseln aus `modelle.py` (`OBJEKT_FELDER`, …). `null` = leer gelassen, die Annahme greift wie in der Mappe; `""` = bewusst leer. Ein Neuobjekt `null` ist eine Annahmezeile (Reinvestition aus dem Verkauf derselben Zeile).
- `afa_plan`: `{ObjektID: {Jahr: AfA}}`; `neukauf`: `[{kst, name, werte: {miete|einnahmen|erhaltung|ausgaben: {Jahr: Wert}}}]`.

**Pfade** der Zahlen im Rechenweg: `par:Name`, `obj:Zeile:Feld`, `prg:b|n:Zeile:Spalte:Jahr`, `vk:Zeile:Spalte`, `rl:Zeile:Spalte`, `rls:Spalte:Jahr`, `ne:Zeile:Spalte`, `neq:Zeile:Quelle:Spalte`, `dl:Zeile:haupt|zweit:Spalte[:Jahr]`, `liq|aus:Szenario:Spalte:Jahr`, `vg:Kennzahl:Szenario`. Zeile und Jahr zählen ab 0.

Wer eine Formel in `prognosemodell/formeln.py` ändert, ändert sie in `rechenwerk.js` mit und prüft mit `python -m pruefung.abgleich`.
