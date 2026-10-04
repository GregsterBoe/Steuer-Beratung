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
| Parameter, Objekte, Verkäufe, Neuobjekte | Prognose, Rücklagen, Liquidität | Auswertung |

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

Stand: Etappen 1 bis 7 umgesetzt (Parameter- und Objektblatt mit Statusprüfung, Prognosematrix mit AfA-Fortschreibung, Indexierung von Miete und Erhaltung, Verkaufsblatt, Rücklagenspiegel nach § 6b, Reinvestition in Neuobjekte, Liquidität und Auswertung).

## Nutzung

```bash
pip install -r requirements.txt
python -m prognosemodell             # erzeugt ausgabe/Prognosemodell_VV.xlsx (mit Testobjekt)
python -m prognosemodell --ohne-testdaten
python -m pruefung.pruefen           # rechnet per LibreOffice headless und prüft gegen Sollwerte
```

Das Prüfskript braucht LibreOffice mit Calc (`soffice`).

Gelb = Eingabe, grau = Formel. Der Schalter Steuerwelt kennt GmbH, Privat/GbR und gewerblich; gerechnet wird im MVP nur GmbH, sonst zeigt das Parameterblatt „nicht im MVP“. Die Statusspalte im Objektblatt meldet fehlende Pflichtfelder, doppelte IDs, ein Kaufjahr nach dem Basisjahr und einen Restbuchwert über den AK. Das Blatt Prognose rechnet je Objekt und Jahr Miete, Erhaltung, AfA, Buchwert und Ergebnis, nur für Objekte mit Status OK. Miete und Erhaltung steigen ab dem Basisjahr mit den Raten vom Parameterblatt. Das Blatt Verkäufe teilt den Erlös eines geplanten Verkaufs in Buchwert und Gewinn, getrennt nach Gebäude und G+B; ab dem Folgejahr rechnet das Objekt nicht mehr mit. Das Blatt Rücklagen prüft je Verkauf die § 6b-Voraussetzungen und bildet die Rücklage getrennt nach Gebäude und G+B. Ein Jahresspiegel zeigt Bildung, Auflösung im Fristjahr samt Gewinnzuschlag, Stand und Steuer auf Veräußerung und Auflösung. Im Blatt Neuobjekte nimmt ein Reinvestitionsobjekt eine Rücklage auf: Die Gebäude-Rücklage geht nur aufs Gebäude, die G+B-Rücklage zuerst auf G+B, der Rest aufs Gebäude. Der Übertrag mindert die AfA-Basis, das Neuobjekt läuft ab dem Folgejahr des Kaufs in der Prognose mit. Die Blätter Liquidität und Auswertung fassen alles je Jahr zusammen: freier Mittelzufluss (Miete minus Erhaltung, Verkaufserlös, Steuer, Kauf von Neuobjekten), Gesamt-GuV, Steuer, Buch- und Verkehrswert sowie stille Reserven, jeweils mit Summenzeile.

Details, Formeln und Testfälle stehen in [docs/Projektplan.md](docs/Projektplan.md).

> Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
