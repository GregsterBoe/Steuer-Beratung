# Prognosemodell V+V

20-Jahres-Prognose (2027–2046, Basis 2026) für Immobilien im Betriebsvermögen einer GmbH: AfA-Fortschreibung, Verkauf, § 6b-Rücklage und Reinvestition. Am Ende werden Szenarien verglichen: 6b-Kette, sofort versteuern und das Kapital anlegen, oder sofort versteuern und trotzdem reinvestieren.

## Prinzipien

- **Endprodukt:** eine eigenständige Excel-Datei, die der Mandant ohne Python bedient.
- **Rechnung in Zellformeln:** jeder Schritt ist im Blatt nachvollziehbar. VBA steuert nur Objekte, Szenarien, Prüfungen und Neuberechnung.
- **Python als Generator:** openpyxl baut die Mappe einmalig beim Erstellen. Ein Prüfskript rechnet sie per LibreOffice headless durch und gleicht sie mit den Testfällen ab.
- **Alle Annahmen auf dem Parameterblatt**, keine festen Zahlen in Formeln.
- **Vor Finanzierung:** Zins und Tilgung kommen erst in Stufe 2.

## Blätter

| Eingabe | Rechnung | Ausgabe | Kontrolle |
| --- | --- | --- | --- |
| Parameter, Objekte, Verkäufe, Neuobjekte | Prognose, Rücklagen, Liquidität | Übersicht, Vergleich, Auswertung | Prüfung, Varianten |

## Etappen

1. Gerüst
2. AfA-Fortschreibung
3. Indexierung
4. Verkauf
5. § 6b-Rücklage
6. Reinvestition
7. Liquidität und Auswertung
8. Szenariovergleich
9. VBA-Steuerung und Plausibilitätsprüfungen

Stand: Etappen 1 bis 9 sind umgesetzt.
- Parameter- und Objektblatt mit Statusprüfung
- Prognoseblatt mit AfA-Fortschreibung und Indexierung je Objekt und Jahr
- Verkaufsblatt mit Aufteilung des Erlöses und Veräußerungsgewinn getrennt nach Gebäude und G+B
- Rücklagenblatt mit § 6b-Rücklage je Verkauf und Spiegel je Jahr
- Neuobjektblatt mit Übertragung der Rücklage und geminderter AfA-Basis
- Liquiditätsblatt mit Steuer, Verlustvortrag und Geldfluss je Jahr, Plan gegen Baseline
- Auswertungsblatt mit Gesamt-GuV, stillen Reserven, latenter Steuer und Gesamtvermögen
- Übersichtsblatt mit Immobilienwert und Gesamtvermögen, je mit Diagramm
- Szenarien A, B, C und Baseline gleichzeitig, mit verzinster Alternativanlage und Vergleichsblatt
- Einleseschicht und Vorlage für die Kostenstellenblätter (DATEV-BWA)
- Prüfungsblatt mit allen Plausibilitätsprüfungen als Formeln, VBA-Steuerung als .xlsm

## Nutzung

```bash
pip install -r requirements.txt
python -m prognosemodell             # erzeugt ausgabe/Prognosemodell_VV.xlsx (mit Testobjekt)
python -m prognosemodell --ohne-testdaten
python -m prognosemodell --makros    # .xlsm mit VBA-Steuerung (braucht LibreOffice beim Bauen)
python -m prognosemodell --kostenstellen Kostenstellen.xlsx --ausgabe Ordner/Prognose.xlsx   # laufende Werte je Blatt einlesen; nur Ordner = Standardname darin
python -m pruefung.pruefen           # rechnet per LibreOffice headless und prüft gegen Sollwerte
python -m pruefung.pruefen "Etappe 9"  # nur Fälle, deren Name den Text enthält
python -m pruefung.pruefen_einlesen  # prüft die Einleseschicht, ohne LibreOffice
python -m prognosemodell.vorlagen    # schreibt vorlagen/Kostenstellen_BWA_Vorlage.xlsx
```

Das Prüfskript und `--makros` brauchen LibreOffice mit Calc und der Python-UNO-Brücke (`soffice`, unter Debian/Ubuntu die Pakete `libreoffice-calc` und `python3-uno`).

Zielformat der Eingabe ist die DATEV-BWA-Kostenstellenblattsammlung; `vorlagen/Kostenstellen_BWA_Vorlage.xlsx` zeigt das Layout mit erfundenen Werten. Als Jahresspalte gilt ein Kopf wie 2026, „Jahr 2026“ oder „Plan 2027“; ein Summenblatt „Alle Objekte“ wird übersprungen. Geplant ist, die Ergebnisse wieder in diese Struktur zu schreiben, mit einem Sonderbereich für Verkauf und Kauf (Projektplan, Abschnitt 18).

`--kostenstellen` überspringt jedes Blatt, das nicht im Kostenstellenformat ist (kein „Nr.“ in B4, keine Kostenstelle in B2 oder keine Spalte des Basisjahrs in Zeile 4), etwa Annahmen oder Übersichten. Jedes übersprungene Blatt nennt es mit Grund in der Ausgabe. Je Kostenstellenblatt liest es B2 (Kostenstelle = ObjektID), C2 (Objektname) und aus der Spalte des Basisjahrs die BWA-Zeilen 1020 (Miete), 1090 (weitere Einnahmen), 1250 (Erhaltung) sowie 1100–1220 und 1260 (weitere Ausgaben). Die Datei muss in Excel gespeichert sein, damit berechnete Werte vorliegen. Steuerliche Stammdaten (AK, Kaufjahr, AfA, Restbuchwert) kommen nicht aus diesen Blättern; solange sie fehlen, meldet die Statusspalte „Pflichtfeld fehlt“.

Das Prognoseblatt hat je Objektzeile einen Block mit 20 Jahreszeilen: Miete, weitere Einnahmen, Erhaltung und weitere Ausgaben wachsen mit ihren Steigerungsraten. Die AfA beträgt AK Gebäude × Satz, höchstens aber den Restbuchwert. Danach sind AfA und Buchwert null.

Das Blatt **Übersicht** öffnet als erstes. Es zeigt von 2026 bis 2046 den Wert der Immobilien und das Gesamtvermögen am Jahresende, je als Tabelle und Liniendiagramm, und zwar in zwei Linien:

- **Baseline:** alles halten, nichts verkaufen.
- **Plan:** mit den Verkäufen und Neuobjekten aus den Blättern Verkäufe und Neuobjekte.

Der Wert ist der Verkehrswert aus dem Objektblatt, fortgeschrieben mit der Wertsteigerung vom Parameterblatt. Objekte ohne Verkehrswert zählen mit 0, die Übersicht zeigt ihre Anzahl rot an. Ein Verkauf gilt zum Jahresende. Miete und AfA laufen im Verkaufsjahr noch, ab dem Folgejahr ist das Objekt inaktiv. Der Plan enthält auch die Neuobjekte. Das Gesamtvermögen ist der Verkehrswert plus die kumulierte Liquidität nach Steuern, abzüglich der latenten Steuer auf stille Reserven und Rücklage. So stehen Halten und Verkaufen vergleichbar nebeneinander.

Im Blatt **Verkäufe** stehen je Verkauf ObjektID, Jahr, Preis, Kosten, optional der Anteil G+B laut Kaufvertrag und § 6b ja/nein. Daraus rechnet das Blatt:
- den Gebäudebuchwert am Ende des Verkaufsjahrs aus der Prognose
- die Aufteilung des Nettoerlöses auf Gebäude und G+B
- den Gewinn je Teil

Die Aufteilung folgt dem Kaufvertrag, sonst dem Verkehrswertanteil aus dem Objektblatt. Der Status meldet unter anderem eine fehlende Aufteilung und eine zu kurze Vorbesitzzeit für § 6b.

Das Blatt **Rücklagen** bildet bei § 6b ja und Status OK eine Rücklage aus den positiven Teilgewinnen, getrennt nach Gebäude und G+B. Die Frist beträgt vier Jahre, mit „§ 6b Neubau begonnen = ja“ sechs Jahre. Was bis zum Fristjahr nicht auf Neuobjekte übertragen ist, wird dort aufgelöst, mit 6 % Zuschlag je Jahr. Der Spiegel je Jahr zeigt Gewinne, Einstellung, Übertragung, Auflösung, Zuschlag, Bestand und den steuerwirksamen Betrag, der in Etappe 7 die Steuer ergibt.

Im Blatt **Neuobjekte** stehen je Reinvestition Kaufjahr, Kaufpreis, Anteil G+B, Nebenkosten, AfA-Satz, AfA-Methode (linear oder degressiv 5 % nach § 7 Abs. 5a EStG mit Wechsel zur linearen AfA), Mietrendite, Erhaltungsquote und die Quelle-Rücklage. Gekauft wird zum Jahresende, Miete und AfA laufen ab dem Folgejahr. Die Rücklage wird in fester Reihenfolge übertragen:
1. Gebäudegewinn auf das neue Gebäude
2. G+B-Gewinn auf den neuen G+B (bis auf 0)
3. Rest des G+B-Gewinns auf das Gebäude

Die AfA läuft von der geminderten AfA-Basis. Neuobjekte erscheinen in der Prognose und in der Plan-Linie der Übersicht, nicht in der Baseline.

Das Blatt **Liquidität** rechnet je Jahr die Steuer und den Geldfluss, je Szenario eine Tabelle:
- Steuer = (laufendes Ergebnis + steuerwirksamer Betrag aus dem Rücklagenspiegel) × Grenzsteuersatz
- Verluste werden vorgetragen und mit späteren Gewinnen verrechnet
- freier Mittelzufluss = Mieten und Einnahmen − Erhaltung und Ausgaben + Zins + Verkaufserlöse − Steuer − Kauf der Neuobjekte

Die Liquidität liegt in einer Alternativanlage und wird mit der Rendite vom Parameterblatt verzinst (Standard 3 %); der Zins ist steuerpflichtig.

Das Blatt **Auswertung** zeigt je Jahr die Gesamt-GuV, Steuer, kumulierte Steuer, stille Reserven, Rücklagenbestand, Verlustvortrag und Gesamtvermögen vor und nach latenter Steuer. Die latente Steuer ist die Steuer, die anfiele, wenn alle Objekte zum Verkehrswert verkauft und die Rücklage aufgelöst würden.

Alle Szenarien rechnen gleichzeitig, Liquidität und Auswertung haben je Szenario eine Tabelle:
- **A Plan:** § 6b-Kette wie erfasst
- **B:** jeder Veräußerungsgewinn sofort versteuert, Neuobjekte mit Quelle-Rücklage entfallen, das Geld bleibt in der Alternativanlage
- **C:** sofort versteuert, die Neuobjekte werden trotzdem gekauft, mit voller AfA-Basis
- **Baseline:** alles halten

Das Blatt **Vergleich** stellt die Kennzahlen am Ende des Rasters nebeneinander, mit den Differenzen A − B, A − C und A − Baseline, dazu das Endvermögen nach latenter Steuer je Jahr als Diagramm. A − C zeigt die reine Wirkung von § 6b: ohne Alternativrendite ein Gleichstand, mit Rendite der Zins auf die gestundete Steuer.

Das Blatt **Prüfung** rechnet alle Plausibilitätsprüfungen als Formeln, je mit Art und Anzahl betroffener Zeilen:
- **Fehler:** Status ungleich OK in Objekten, Verkäufen oder Neuobjekten (darunter Fristverstoß), Steuerwelt nicht GmbH
- **Warnung:** Vorbesitzzeit für § 6b zu kurz, Rücklage nur teilweise oder gar nicht übertragen, Drei-Objekt-Grenze überschritten
- **Hinweis:** Frist nach Prognoseende, Liquidität negativ (Finanzierungsbedarf), Verkehrswert fehlt

Das Gesamtergebnis steht auch auf dem Parameterblatt und in der Übersicht.

Mit `--makros` entsteht eine .xlsm mit VBA-Steuerung, die Schaltflächen liegen auf dem Parameterblatt:
- Plausibilität prüfen
- Neu berechnen
- Objekt anlegen, duplizieren, entfernen
- leere Prognoseblöcke aus- und einblenden
- Variante festhalten

„Variante festhalten“ schreibt die Kennzahlen des Vergleichs als feste Werte ins Blatt **Varianten**, so lassen sich etwa verschiedene Verkaufsjahre vergleichen. VBA rechnet nichts, alle Ergebnisse entstehen in den Formeln.

Gelb = Eingabe, grau = Formel. Die Statusspalte im Objektblatt meldet fehlende Pflichtfelder, doppelte IDs, ein Kaufjahr nach dem Basisjahr und einen Restbuchwert über den AK.

Details, Formeln und Testfälle stehen in [docs/Projektplan.md](docs/Projektplan.md).

> Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
