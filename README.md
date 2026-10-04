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

Details, Formeln und Testfälle stehen in [docs/Projektplan.md](docs/Projektplan.md).

> Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
