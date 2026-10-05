# Prognosemodell V+V

20-Jahres-Prognose (2027–2046, Basis 2026) für Immobilien im Betriebsvermögen einer GmbH: AfA-Fortschreibung, Verkauf, § 6b-Rücklage und Reinvestition. Am Ende werden zwei Szenarien verglichen: 6b-Kette oder sofort versteuern und das Kapital anderweitig anlegen.

## Prinzipien

- **Endprodukt:** eine eigenständige Excel-Datei, die der Mandant ohne Python bedient.
- **Rechnung in Zellformeln:** jeder Schritt ist im Blatt nachvollziehbar. VBA steuert nur Objekte, Szenarien, Prüfungen und Neuberechnung.
- **Python als Generator:** openpyxl baut die Mappe einmalig beim Erstellen. Ein Prüfskript rechnet sie per LibreOffice headless durch und gleicht sie mit den Testfällen ab.
- **Alle Annahmen auf dem Parameterblatt**, keine festen Zahlen in Formeln.
- **Vor Finanzierung:** Zins und Tilgung kommen erst in Stufe 2.

## Blätter

| Eingabe | Rechnung | Ausgabe |
| --- | --- | --- |
| Parameter, Objekte, Verkäufe, Neuobjekte | Prognose, Rücklagen, Liquidität | Übersicht, Auswertung |

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

Stand: Etappen 1 bis 7 sind umgesetzt.
- Parameter- und Objektblatt mit Statusprüfung
- Prognoseblatt mit AfA-Fortschreibung und Indexierung je Objekt und Jahr
- Verkaufsblatt mit Aufteilung des Erlöses und Veräußerungsgewinn getrennt nach Gebäude und G+B
- Rücklagenblatt mit § 6b-Rücklage je Verkauf und Spiegel je Jahr
- Neuobjektblatt mit Übertragung der Rücklage und geminderter AfA-Basis
- Liquiditätsblatt mit Steuer, Verlustvortrag und Geldfluss je Jahr, Plan gegen Baseline
- Auswertungsblatt mit Gesamt-GuV, stillen Reserven, latenter Steuer und Gesamtvermögen
- Übersichtsblatt mit Immobilienwert und Gesamtvermögen, je mit Diagramm
- Einleseschicht für die Kostenstellenblätter (DATEV-BWA)

## Nutzung

```bash
pip install -r requirements.txt
python -m prognosemodell             # erzeugt ausgabe/Prognosemodell_VV.xlsx (mit Testobjekt)
python -m prognosemodell --ohne-testdaten
python -m prognosemodell --kostenstellen Kostenstellen.xlsx --ausgabe Ordner/Prognose.xlsx   # laufende Werte je Blatt einlesen; nur Ordner = Standardname darin
python -m pruefung.pruefen           # rechnet per LibreOffice headless und prüft gegen Sollwerte
python -m pruefung.pruefen_einlesen  # prüft die Einleseschicht, ohne LibreOffice
```

Das Prüfskript braucht LibreOffice mit Calc (`soffice`, unter Debian/Ubuntu Paket `libreoffice-calc`).

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

Im Blatt **Neuobjekte** stehen je Reinvestition Kaufjahr, Kaufpreis, Anteil G+B, Nebenkosten, AfA-Satz, Mietrendite, Erhaltungsquote und die Quelle-Rücklage. Gekauft wird zum Jahresende, Miete und AfA laufen ab dem Folgejahr. Die Rücklage wird in fester Reihenfolge übertragen:
1. Gebäudegewinn auf das neue Gebäude
2. G+B-Gewinn auf den neuen G+B (bis auf 0)
3. Rest des G+B-Gewinns auf das Gebäude

Die AfA läuft von der geminderten AfA-Basis. Neuobjekte erscheinen in der Prognose und in der Plan-Linie der Übersicht, nicht in der Baseline.

Das Blatt **Liquidität** rechnet je Jahr die Steuer und den Geldfluss, links für den Plan, rechts für die Baseline:
- Steuer = (laufendes Ergebnis + steuerwirksamer Betrag aus dem Rücklagenspiegel) × Grenzsteuersatz
- Verluste werden vorgetragen und mit späteren Gewinnen verrechnet
- freier Mittelzufluss = Mieten und Einnahmen − Erhaltung und Ausgaben + Verkaufserlöse − Steuer − Kauf der Neuobjekte

Die Liquidität wird nicht verzinst; die Alternativanlage kommt mit Etappe 8.

Das Blatt **Auswertung** zeigt je Jahr die Gesamt-GuV, Steuer, kumulierte Steuer, stille Reserven, Rücklagenbestand, Verlustvortrag und Gesamtvermögen vor und nach latenter Steuer. Die latente Steuer ist die Steuer, die anfiele, wenn alle Objekte zum Verkehrswert verkauft und die Rücklage aufgelöst würden.

Gelb = Eingabe, grau = Formel. Die Statusspalte im Objektblatt meldet fehlende Pflichtfelder, doppelte IDs, ein Kaufjahr nach dem Basisjahr und einen Restbuchwert über den AK.

Details, Formeln und Testfälle stehen in [docs/Projektplan.md](docs/Projektplan.md).

> Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
