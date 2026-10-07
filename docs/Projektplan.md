# Prognosemodell V+V – Entwicklungsplan Excel/VBA (MVP)

Oct 4, 2026 · @Gregor

MVP mit voller Logik inklusive § 6b-Rücklagen; Rechenlogik in Formeln, VBA nur zur Steuerung, Finanzierung in einer späteren Stufe.

## 1. Umfang und Prinzipien

Der MVP bildet die volle steuerliche Logik ab: Abschreibungsdynamik, Verkauf, § 6b-Rücklage und Reinvestition über 20 Jahre, für eine gewerbliche Struktur mit GmbH. Finanzierung bleibt zunächst außen vor.

**Prinzipien**

- Rechenlogik liegt in Zellformeln, nicht in VBA. So ist jeder Schritt im Blatt nachvollziehbar und prüfbar.
- VBA steuert nur: Objekte anlegen und ausblenden, Szenarien speichern und vergleichen, Plausibilitätsprüfungen, Neuberechnung.
- Eine Zeile je Objekt und Jahr, keine verbundenen Zellen, keine festen Zahlen in Formeln. Alle Annahmen stehen auf dem Parameterblatt.
- Jahresraster 2027 bis 2046, Basis Ist 2026. Verkauf und Kauf zum Jahresende.
- Jede Auswertung ist als "vor Finanzierung" gekennzeichnet, solange Zins und Tilgung fehlen.

**Nicht im MVP:** Zins, Tilgung, Restschuld, Vorfälligkeit; Sensitivitäten; IRR; Steuer auf Ausschüttungsebene. Diese Punkte sind in der Situationsbeschreibung als spätere Stufe vermerkt.

## 2. Blattarchitektur

Der Datenfluss läuft in eine Richtung: Eingabeblätter speisen die Rechenblätter, diese die Auswertung. Rückbezüge gibt es nur innerhalb der Rechnung, etwa wenn eine Reinvestition die Prognose eines Neuobjekts anstößt.

```text
Eingabe:  Parameter, Objekte, Anlagen, Verkäufe, Neuobjekte
            ↓
Rechnung: Prognose ⇄ Rücklagen → Liquidität
            ↓
Ausgabe:  Auswertung, Vergleich, Übersicht
            ↓
Kontrolle: Prüfung, Varianten (Etappe 9)
            ↓
BWA:      Alle Objekte, je Kostenstelle ein Blatt, Verkauf und Kauf (Abschnitt 18)
```

So bleibt nachvollziehbar, woher jede Zahl kommt: Alle Eingaben links, die Rechnung in der Mitte, die Ergebnisse rechts.

## 3. Blätter im Detail

Dreizehn feste Blätter, getrennt nach Eingabe, Rechnung, Ausgabe und Kontrolle. Dazu kommen die BWA-Ausgabe (Summenblatt, je Kostenstelle ein Blatt) und der Sonderbereich Verkauf und Kauf (Abschnitt 18). Eingabeblätter sind die einzige Stelle, an der getippt wird.

| Blatt | Typ | Schlüsselfelder | Zweck |
| --- | --- | --- | --- |
| Start | Ausgabe | Handlungsempfehlung, Endvermögen je Option, Datenlage, Anleitung, Farblegende, zentrale Annahmen | erstes Blatt, Einstieg und Ergebnis auf einen Blick (Abschnitt 20) |
| Parameter | Eingabe | Steuerwelt-Schalter, Grenzsteuersatz, Mietsteigerung, Erhaltungssteigerung, Kostensteigerung, Wertsteigerung, GrESt-Satz, degressive AfA, Alternativrendite, § 6b-Fristen | alle globalen Annahmen und Schalter |
| Objekte | Eingabe | ObjektID, AK Gebäude, AK G+B, AfA-Satz, Kaufjahr, Restbuchwert 2026, Miete 2026, Erhaltung 2026 | Stammdaten je Bestandsobjekt |
| Anlagen | Eingabe | Inventar-Nr., ObjektID, Art, AHK, Buchwert Stand, AfA-Methode, AfA-Satz, AfA p. a.; AfA je Prognosejahr | Anlagenverzeichnis, liefert AK, Buchwert, Kaufjahr und AfA je Objekt (Abschnitt 21) |
| Verkäufe | Eingabe | ObjektID, Verkaufsjahr, Verkaufspreis oder Faktor, Verkaufskosten, 6b-Nutzung (ja/nein), Neubau begonnen (ja/nein) | ein Datensatz je geplantem Verkauf |
| Neuobjekte | Eingabe | NeuID, Kaufjahr, Kaufpreis, Anteil G+B, Nebenkosten, AfA-Satz, AfA-Methode, Mietrendite, Erhaltungsquote, Quelle-Rücklage | Reinvestitionsobjekte |
| Prognose | Rechnung | ObjektID × Jahr (2027-2046), Miete, Erhaltung, AfA, Buchwert, Ergebnis, aktiv-Flag | Jahresmatrix je Objekt, Herzstück |
| Rücklagen | Rechnung | RücklageID, Verkaufsjahr, Betrag G+B, Betrag Gebäude, Fristjahr, übertragen Gebäude und G+B, Auflösung, Zuschlag; Spiegel je Jahr | § 6b-Spiegel je Rücklage und je Jahr |
| Liquidität | Rechnung | Jahr, Steuer mit Verlustvortrag, Verkaufserlöse, Buchwert-Rückfluss, Kauf Neuobjekte, Liquidität kumuliert, Zinsertrag; Szenarien A, B, C und Baseline | Steuer und Geldfluss je Jahr |
| Auswertung | Ausgabe | Jahr, Gesamt-GuV, Steuer, stille Reserven, latente Steuer, Gesamtvermögen; Szenarien A, B, C und Baseline | Kennzahlen je Jahr |
| Vergleich | Ausgabe | Kennzahlen je Szenario am Ende des Rasters, Differenzen A − B, A − C, A − Baseline; Endvermögen je Jahr mit Diagramm | Entscheidung § 6b-Kette oder sofort versteuern |
| Übersicht | Ausgabe | Jahr, Verkehrswert und Gesamtvermögen je Baseline und Plan, Differenzen, zwei Diagramme | Immobilienwert und Gesamtvermögen im Jahresverlauf, Plan mit Verkäufen und Neuobjekten gegen Nichtstun |
| Prüfung | Kontrolle | je Plausibilitätsprüfung Art, Anzahl betroffener Zeilen, Ergebnis; Gesamtergebnis | alle Plausibilitätsprüfungen als Formeln (Abschnitt 19) |
| Alle Objekte, je Kostenstelle ein Blatt | Ausgabe | BWA-Zeilen 1010–1380, Ist links, Planjahre rechts als Formeln | Ergebnisse im DATEV-BWA-Format (Abschnitt 18) |
| Verkauf und Kauf | Ausgabe | je Verkauf Ergebnis- und Detailsicht, je Neuobjekt Detailsicht Kauf | Sonderbereich für die Berichterstattung (Abschnitt 18) |
| Varianten | Kontrolle | je festgehaltener Variante Bezeichnung, Zeitpunkt, Endvermögen A, B, C, Baseline, Differenzen, Steuer | Ergebnisse verschiedener Eingaben vergleichen; nur das Makro schreibt hier |

Konvention: ObjektID ist der Schlüssel, der Objekte, Verkäufe, Rücklagen und Prognose verbindet. Neuobjekte bekommen eine eigene ID, laufen in der Prognose aber in derselben Matrix.

## 4. Rechenkern in Formeln

Vier Rechenbausteine, alle als Zellformeln. Die Notation unten ist fachlich, in Excel werden daraus SUMMEWENNS-, WENN- und MAX-Formeln über die Jahresspalten.

**AfA-Fortschreibung je Objekt und Jahr t**

- Miete\_t = Miete\_2026 × (1 + Mietsteigerung)^(t − 2026)
- Erhaltung\_t = Erhaltung\_2026 × (1 + Erhaltungssteigerung)^(t − 2026)
- AfA\_t = MIN(AK\_Gebäude × AfA-Satz; Restbuchwert\_(t−1)), also keine AfA mehr, wenn der Buchwert null ist
- Buchwert\_t = Restbuchwert\_(t−1) − AfA\_t
- Ergebnis\_t = (Miete\_t + weitere Einnahmen\_t − Erhaltung\_t − weitere Ausgaben\_t − AfA\_t) × aktiv-Flag
- weitere Einnahmen wachsen mit der Mietsteigerung, weitere Ausgaben mit par\_Kostensteig
- aktiv-Flag = 1, solange t ≤ Verkaufsjahr, sonst 0

**Verkaufsaufteilung im Verkaufsjahr**

- Restbuchwert gesamt = Buchwert Gebäude + AK G+B
- Veräußerungsgewinn = Verkaufspreis − Verkaufskosten − Restbuchwert gesamt
- Gewinn Gebäude = Erlösanteil Gebäude − Buchwert Gebäude; Gewinn G+B = Erlösanteil G+B − AK G+B
- Erlösanteile werden nach dem Anteil G+B laut Kaufvertrag aufgeteilt, ersatzweise nach dem Verkehrswertanteil aus dem Objektblatt (keine Aufteilung nach Buchwerten, siehe Abschnitt 11)

**Rücklagenspiegel § 6b**

- Bei 6b-Nutzung = ja und Vorbesitzzeit ≥ 6 Jahre: Rücklage = Veräußerungsgewinn, getrennt nach G+B und Gebäude
- Nur positive Teilgewinne gehen in die Rücklage; ein Verlust des anderen Teils wirkt sofort
- Fristjahr = Verkaufsjahr + 4, bei begonnenem Neubau + 6 (§ 6b Abs. 3)
- Übertrag im Kaufjahr eines Neuobjekts: Gebäudegewinn nur auf Gebäudeanteil, G+B-Gewinn auf beides (§ 6b Abs. 1 EStG)
- Reihenfolge: Gebäudegewinn auf neues Gebäude, G+B-Gewinn zuerst auf neuen G+B (kostet keine AfA), nur der Rest auf das Gebäude (Abschnitt 13)
- Neue AfA-Basis Gebäude = Gebäude-AK Neuobjekt − übertragener Gebäudegewinn
- Nicht genutzt bis Fristjahr: Auflösung + 6 % je vollem Jahr als Ertrag (4 Jahre 24 %, Neubau 36 %)

**Szenariovergleich**

- Szenario A (6b-Kette): Steuer im Verkaufsjahr = 0 bei voller Rücklage; dafür spätere AfA geringer
- Szenario B (sofort versteuern): Steuer = Veräußerungsgewinn × Grenzsteuersatz; freigesetztes Kapital in Alternativanlage mit eigener Rendite
- Szenario C (sofort versteuern, trotzdem kaufen): wie B, aber die Neuobjekte werden gekauft, mit voller AfA-Basis
- Vergleich über das Endvermögen nach latenter Steuer nach 20 Jahren; alle Szenarien rechnen gleichzeitig (Abschnitt 15)

## 5. VBA-Module

VBA steuert nur, es rechnet nicht. Die Makros schreiben Eingabewerte und lösen Neuberechnung aus; die Ergebnisse entstehen in den Formeln.

| Modul | Aufgabe |
| --- | --- |
| modObjekte | Objekt anlegen, duplizieren, ausblenden; Prognosezeilen je Objekt erzeugen |
| modSzenario | entfällt: alle Szenarien rechnen gleichzeitig in Formeln (Abschnitt 15) |
| modVarianten | Kennzahlen des Blatts Vergleich als feste Werte im Blatt Varianten festhalten |
| modPruefung | Plausibilitätsprüfungen vor der Rechnung (siehe unten) |
| modRechnen | Neuberechnung anstoßen, Auswertung aktualisieren |
| modStart | Menü bzw. Schaltflächen auf dem Parameterblatt |

**Plausibilitätsprüfungen in modPruefung**

- Reinvestitionsjahr ≤ Verkaufsjahr + 4, sonst Hinweis auf Fristverstoß
- Gebäudeanteil des Neuobjekts ≥ übertragener Gebäudegewinn
- Vorbesitzzeit des verkauften Objekts ≥ 6 Jahre
- Summe der Objektanteile G+B plus Gebäude = Kaufpreis
- mehr als 3 Verkäufe in 5 Jahren: Warnung wegen gewerblichem Grundstückshandel
- fehlende Pflichtfelder je aktivem Objekt

Prüfungen schreiben ihr Ergebnis in eine Statusspalte, nicht in Pop-ups allein, damit Fehler im Blatt sichtbar bleiben. Umgesetzt ist das in Etappe 9 als Blatt Prüfung (Abschnitt 19).

## 6. Umsetzungsreihenfolge

In Etappen, jede mit prüfbarem Zwischenstand. Erst wenn eine Etappe an einem Objekt stimmt, kommt die nächste.

1. **Gerüst:** Parameter- und Objektblatt anlegen, ein Testobjekt erfassen. Prüfbar: Stammdaten vollständig.
2. **AfA-Fortschreibung:** Prognosematrix für ein Objekt über 20 Jahre, ohne Verkauf. Prüfbar: Buchwert läuft korrekt auf null, AfA stoppt danach.
3. **Indexierung:** Miete und Erhaltung mit Steigerungsraten. Prüfbar: Werte wachsen wie erwartet.
4. **Verkauf:** Verkaufsblatt, Aufteilung in Buchwert und Gewinn, Objekt ab Folgejahr inaktiv. Prüfbar: Gewinn = Preis minus Buchwert minus Kosten.
5. **Rücklage:** Rücklagenspiegel, Bildung und Fristjahr. Prüfbar: Gewinn landet getrennt nach G+B und Gebäude in der Rücklage.
6. **Reinvestition:** Neuobjekt, Übertrag, geminderte AfA-Basis. Prüfbar: neue AfA-Basis stimmt, Gebäudeanteil reicht.
7. **Liquidität und Auswertung:** Geldfluss und Gesamt-GuV. Prüfbar: Summen über alle Objekte.
8. **Szenariovergleich:** A gegen B und C über 20 Jahre. Prüfbar: alle Pfade nachvollziehbar.
9. **VBA-Steuerung und Prüfungen:** erst wenn die Formeln stehen. Prüfbar: Objekt anlegen ohne Formelbruch (Abschnitt 19).

Dazu kommt die Ausgabe im DATEV-BWA-Format mit dem Sonderbereich Verkauf und Kauf (Abschnitt 18). Sie baut auf Etappe 8 auf.

Erst nach Etappe 9 folgt die zweite Stufe mit der Finanzierung. Vorrang hat dort die Restschuld der verkauften Objekte und das Darlehen für den Teil der Reinvestition, den der Erlös nicht deckt (Abschnitt 18).

## 7. Testfälle und Abnahme

Jeder Testfall ist eine Rechnung von Hand, gegen die das Blatt geprüft wird.

| Fall | Eingabe | Erwartetes Ergebnis |
| --- | --- | --- |
| AfA-Ende | Gebäude 800.000, 2,5 %, Kauf 2007, Restbuchwert 2026 400.000 | Buchwert Ende 2046 null, AfA danach 0 |
| AfA läuft im Raster aus | Gebäude 800.000, 2 %, Restbuchwert 2026 50.000 | AfA 2030 nur 2.000, ab 2031 AfA und Buchwert 0, Miete läuft weiter |
| Verkauf mit Gewinn | Preis 1,4 Mio, Buchwert 680.000, Kosten 0, hälftig | Gewinn 720.000, davon Gebäude 220.000, G+B 500.000 |
| Rücklage voll | 6b ja, Gewinn 720.000 | Steuer im Verkaufsjahr 0, Rücklage 720.000 |
| Verkauf mit Kaufvertragsaufteilung | Preis 1,4 Mio, Kosten 40.000, 30 % G+B, Verkauf Ende 2030 | Gewinn Gebäude 536.000, G+B 208.000 |
| Übertrag | Rücklage Gebäude 220.000, G+B 500.000; Neuobjekt G+B 360.000, Gebäude 840.000 | AfA-Basis 480.000, AK G+B 0 (G+B-Gewinn zuerst auf G+B) |
| Frist verpasst | keine Reinvestition bis Fristjahr | Auflösung + 6 % je Jahr |
| Szenariovergleich | A, B, C, gleiche Objekte | Endvermögen je Szenario; ohne Alternativrendite A = C |

**Abnahmekriterien**

- Alle sechs Testfälle stimmen auf den Euro genau mit der Handrechnung.
- Jede Zelle ist eine Formel, keine festen Zahlen außerhalb der Eingabeblätter.
- Plausibilitätsprüfungen melden jeden eingebauten Fehler.
- Jede Auswertung trägt den Hinweis "vor Finanzierung".

Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.

## 8. Prognosematrix (Etappe 2)

Eine Zeile je Objekt und Jahr, das sogenannte Long-Format. Es ist mit SUMMEWENNS und Pivot am leichtesten auszuwerten und wächst sauber mit neuen Objekten.

**Spaltenlayout Blatt Prognose**

| Spalte | Feld | Quelle |
| --- | --- | --- |
| A | ObjektID | aus Objekte/Neuobjekte |
| B | Jahr | 2027 bis 2046 |
| C | aktiv-Flag | Formel |
| D | Miete | Formel |
| E | weitere Einnahmen | Formel |
| F | Erhaltung | Formel |
| G | weitere Ausgaben | Formel |
| H | AfA | Formel |
| I | Buchwert Gebäude Ende | Formel |
| J | Ergebnis vor Finanzierung | Formel |

Zusätzlich für Übersicht und Auswertung: K Verkehrswert Ende (prg\_Verkehrswert, Verkehrswert aktuell × (1 + par\_Wertsteig)^(Jahr − Basisjahr), unabhängig vom Verkauf) und L im Bestand Ende (prg\_Bestand, 1 solange Jahr < Verkaufsjahr). Seit Etappe 6 und 7 folgen M Neuobjekt (prg\_Neu), N Buchwert G+B Ende (prg\_BuchwertGuB, AK G+B, beim Neuobjekt die steuerlichen AK G+B ab dem Kaufjahr) und O Buchwert Gebäude bei Halten (prg\_BuchwertHalten = MAX(Restbuchwert − (Jahr − Basisjahr) × AK Gebäude × AfA-Satz; 0), für Neuobjekte 0). Für den Szenariovergleich (Etappe 8) folgen P Neuobjekt mit Rücklage (prg\_MitQuelle) sowie Q bis S AfA, Gebäudebuchwert und Buchwert G+B ohne § 6b (prg\_AfAOhne6b, prg\_BuchwertOhne6b, prg\_BuchwertGuBOhne6b; Abschnitt 15).

Die Spalten tragen benannte Bereiche (prg\_ID, prg\_Jahr, prg\_Aktiv, prg\_Miete, prg\_Einnahmen, prg\_Erhaltung, prg\_Ausgaben, prg\_AfA, prg\_Buchwert, prg\_Ergebnis) für die SUMMEWENNS der späteren Etappen.

**Umsetzung (Etappe 2):** Jede Zeile des Objektblatts hat im Prognoseblatt einen festen Block mit einer Zeile je Prognosejahr (200 Objektzeilen × 20 Jahre). Die Stammdaten kommen per INDEX(obj\_…; n) direkt aus Objektzeile n statt per SVERWEIS über die ObjektID; so greift auch bei doppelter ID jeder Block auf seine eigene Zeile, und die Statusspalte meldet die Dopplung. Der Buchwert des Vorjahres ist die Zeile darüber, im ersten Jahr der Restbuchwert aus dem Objektblatt; die SUMMEWENNS-Variante unten ist damit nicht nötig. Leere Objektzeilen ergeben leere Prognosezeilen. Das aktiv-Flag liest das Verkaufsjahr per INDEX/MATCH aus dem Blatt Verkäufe; ohne Verkauf ist es 1. Die Indexierung aus Etappe 3 ist bereits enthalten. Das Raster hat fest par\_Prognosejahre Zeilen je Objekt, wie beim Generieren eingestellt.

Die Stammdaten stehen auf dem Blatt Objekte, Suche über die ObjektID in Spalte A. Annahme für die Beispielformeln: Objekte-Spalten sind benannte Bereiche (obj\_ID, obj\_MieteBasis, obj\_AfASatz, obj\_AKGebaeude, obj\_Kaufjahr), und auf dem Parameterblatt stehen par\_Mietsteig, par\_Erhaltsteig sowie par\_Basisjahr (2026).

**Formeln je Zeile (Beispiel Zeile 2)**

Die erste Formel ist das aktiv-Flag; es schaltet ein Objekt ab dem Jahr nach dem Verkauf aus.

```text
C2  =WENN(B2<=WENNFEHLER(SVERWEIS(A2;Verkaeufe!A:B;2;FALSCH);9999);1;0)
```

Miete und Erhaltung wachsen vom Basisjahr 2026 mit ihrer jeweiligen Rate.

```text
D2  =SVERWEIS(A2;obj_Basis;SPALTE_Miete;FALSCH)*(1+par_Mietsteig)^(B2-par_Basisjahr)*C2
E2  =SVERWEIS(A2;obj_Basis;SPALTE_Erh;FALSCH)*(1+par_Erhaltsteig)^(B2-par_Basisjahr)*C2
```

Die AfA stoppt, sobald der Restbuchwert null ist. Dafür braucht es den Buchwert des Vorjahres. Für das erste Prognosejahr ist das der Restbuchwert 2026 aus dem Objektblatt, danach der Wert aus der Zeile des Vorjahres.

```text
F2  =MIN(SVERWEIS(A2;obj_Basis;SPALTE_AKGeb;FALSCH)*SVERWEIS(A2;obj_Basis;SPALTE_AfASatz;FALSCH); Buchwert_Vorjahr)*C2
G2  =MAX(Buchwert_Vorjahr - F2; 0)
H2  =(D2 - E2 - F2)*C2
```

Für Buchwert\_Vorjahr empfehle ich einen SUMMEWENNS-Verweis auf dieselbe Matrix: Buchwert der Zeile mit gleicher ObjektID und Jahr gleich B2 minus 1. Im ersten Jahr greift stattdessen der Restbuchwert aus dem Objektblatt. Das lässt sich mit einem WENN(B2=Startjahr; …; SUMMEWENNS(…)) lösen.

**Abnahme Etappe 2:** Bei Gebäude 800.000 und AfA 2,5 % ab Kauf 2007 (Restbuchwert 2026: 400.000) muss der Buchwert Ende 2046 null erreichen. Läuft der Buchwert im Raster aus, gilt: AfA im letzten Jahr nur noch der Rest, danach sind AfA und Buchwert null, Miete und Erhaltung laufen weiter.

Korrektur: Die ursprüngliche Vorgabe (2 %, Kauf 2007, Buchwert 2047 null) geht nicht auf. 800.000 × 2 % = 16.000 je Jahr reicht 50 Jahre, der Buchwert erreicht null also erst Ende 2056. Ende 2046 stehen noch 160.000. Beide Fälle prüft `pruefung/pruefen.py`.

## 9. Generierung per Python, Endprodukt autarke Excel-Datei

Das Endprodukt ist eine eigenständige Excel-Datei, die ohne Python läuft und vom Mandanten selbst bedient wird. Python ist nur ein einmaliger Generator beim Bauen, kein Teil des laufenden Modells. Die gesamte Rechenlogik liegt in Zellformeln; Python schreibt sie nur hinein.

**Vorgehen**

- openpyxl erzeugt einmalig die Arbeitsmappe: ein Blatt je Abschnitt aus Punkt 3, Kopfzeilen, benannte Bereiche, Datenvalidierung und Dropdowns für die Eingabeblätter. Das Skript läuft auf deinem Rechner oder über Claude Code, nicht in der Mandantenumgebung.
- Die Formeln werden als Strings in die Zellen geschrieben. openpyxl legt sie als echte Excel-Formeln ab; gerechnet wird ausschließlich in Excel beim Öffnen. Die fertige xlsx-Datei enthält keinerlei Python-Abhängigkeit.
- Die Jahresmatrix wird im Skript über eine Schleife je Objekt und Jahr aufgebaut, nicht Zelle für Zelle von Hand.
- Formatierung, Zahlenformate und Blattschutz der Rechenblätter setzt das Skript gleich mit.

**Architekturempfehlung**

- Entkopplung wie bei Pillar 2: ein Modul für die Datenmodelle, eines zum Schreiben der Mappe, eines für die Formel-Bausteine. So bleibt das Blattlayout vom Formelcode getrennt.
- Die Formeln aus Abschnitt 4 und 8 als zentrale Bausteine pflegen, damit eine Änderung an einer Stelle greift.
- Testfälle aus Abschnitt 7 als Prüfskript: Mappe generieren, mit einer Bibliothek wie formulas oder LibreOffice-Headless durchrechnen, Ergebnisse gegen die Handrechnung prüfen.

**Grenze zu VBA:** Die VBA-Steuerung aus Abschnitt 5 ist davon unberührt. Sie wird als Makromodul mitgeliefert und kann ebenfalls per Skript in die Mappe eingebettet werden, sobald die Formel-Mappe steht.

## 10. Prüfskript gegen die Testfälle

Das Prüfskript stellt sicher, dass die generierte Mappe korrekt rechnet. Es läuft nur bei dir beim Bauen, nicht beim Mandanten. Der Ablauf: Mappe generieren, mit Formelwerten durchrechnen, Ergebnisse gegen die Handrechnung aus Abschnitt 7 prüfen.

**Das Problem mit den Formelwerten**

openpyxl schreibt Formeln, berechnet sie aber nicht. Eine frisch generierte Datei enthält die Formeln, aber noch keine Ergebniswerte. Zum Prüfen muss also etwas die Mappe einmal durchrechnen.

**Zwei gangbare Wege**

- **LibreOffice headless** (empfohlen): LibreOffice öffnet die Datei per Kommandozeile, rechnet alle Formeln und speichert sie neu. Danach liest openpyxl mit data\_only=True die echten Ergebniswerte. Vorteil: rechnet Excel-Formeln sehr originalgetreu, auch SVERWEIS und SUMMEWENNS. Nachteil: LibreOffice muss auf dem Baurechner installiert sein.
- **Die Bibliothek formulas**: wertet Excel-Formeln direkt in Python aus, ohne Office. Vorteil: leichtgewichtig. Nachteil: deckt nicht jede Funktion gleich gut ab, bei Randfällen muss man gegenprüfen.

**Ablauf des Skripts**

1. Testmappe mit bekannten Eingaben generieren, eine je Testfall aus Abschnitt 7.
2. Mappe durchrechnen lassen (LibreOffice headless oder formulas).
3. Zielzellen auslesen, etwa Buchwert 2047 oder Veräußerungsgewinn.
4. Mit der erwarteten Zahl vergleichen, auf den Euro genau, Toleranz ein Cent für Rundung.
5. Bei Abweichung: Fall, Zelle, Soll und Ist ausgeben und das Skript mit Fehlercode beenden.

**Nutzen**

So wird jede Änderung an den Formel-Bausteinen automatisch gegen alle sechs Fälle geprüft, bevor die Mappe an den Mandanten geht. Das Skript ist das Sicherheitsnetz zwischen Formeländerung und Auslieferung.

## 11. Verkauf (Etappe 4)

Im Verkaufsjahr wird der Erlös in den steuerneutralen Buchwert-Rückfluss und den steuerpflichtigen Gewinn zerlegt, getrennt nach Gebäude und Grund und Boden. Verkauft wird zum Jahresende: Miete und AfA laufen im Verkaufsjahr noch, ab dem Folgejahr ist das Objekt inaktiv.

**Spaltenlayout Blatt Verkäufe (umgesetzt)**

| Spalte | Feld | Name | Quelle |
| --- | --- | --- | --- |
| A | ObjektID | vk\_ID | Eingabe, Auswahl aus Objekte |
| B | Verkaufsjahr | vk\_Jahr | Eingabe |
| C | Verkaufspreis | vk\_Preis | Eingabe |
| D | Verkaufskosten | vk\_Kosten | Eingabe, leer = 0 |
| E | Anteil G+B lt. Kaufvertrag | vk\_AnteilGuBVertrag | Eingabe, optional |
| F | § 6b nutzen | vk\_6b | Dropdown ja/nein |
| G | § 6b Neubau begonnen | vk\_6bNeubau | Dropdown ja/nein, verlängert die Frist (Abschnitt 12) |
| H | Vorbesitzzeit Jahre | vk\_Vorbesitz | Verkaufsjahr − Kaufjahr |
| I | Buchwert Gebäude Ende Verkaufsjahr | vk\_BuchwertGeb | SUMIFS über prg\_Buchwert |
| J | AK G+B | vk\_AKGuB | aus Objekte |
| K | Nettoerlös | vk\_Nettoerloes | C − D |
| L | Anteil G+B verwendet | vk\_AnteilGuB | E, sonst 1 − Verkehrswertanteil Gebäude |
| M | Erlösanteil Gebäude | vk\_ErloesGeb | K − N |
| N | Erlösanteil G+B | vk\_ErloesGuB | K × L |
| O | Gewinn Gebäude | vk\_GewinnGeb | M − I |
| P | Gewinn G+B | vk\_GewinnGuB | N − J, G+B wird nicht abgeschrieben |
| Q | Veräußerungsgewinn | vk\_Gewinn | O + P |
| R | Status | vk\_Status | Plausibilität |

**Formeln (Zeile 2, englische Syntax wie in der Mappe)**

```text
I2  =SUMIFS(prg_Buchwert, prg_ID, A2, prg_Jahr, B2)
L2  =IF(E2<>"", E2, IF(INDEX(obj_VKQuoteGeb, MATCH(A2,obj_ID,0))="", "", 1-INDEX(obj_VKQuoteGeb, MATCH(A2,obj_ID,0))))
N2  =K2*L2
O2  =M2-I2
P2  =N2-J2
```

Die berechneten Spalten bleiben leer, solange die ObjektID fehlt oder unbekannt ist. Der Buchwert in I ist der Prognosewert am Ende des Verkaufsjahrs, die AfA des Verkaufsjahrs ist also schon abgezogen.

**Korrektur gegenüber dem ersten Entwurf:** Die Aufteilung nach dem Verhältnis der Buchwerte (früher Parameter par\_Aufteilung) ist entfallen. Steuerlich maßgeblich ist die Aufteilung im Kaufvertrag, solange sie die realen Wertverhältnisse nicht verfehlt, sonst das Verhältnis der Verkehrswerte (BMF-Arbeitshilfe). Die Buchwertmethode hätte den Gewinn proportional zu den Buchwerten verteilt und den G+B-Anteil systematisch zu niedrig angesetzt. Der G+B-Anteil ist aber der flexibel übertragbare Teil (Abschnitt 13), deshalb zählt die Aufteilung für die Optimierung.

**Status je Zeile**, in dieser Reihenfolge:
- ObjektID unbekannt
- Objekt mehrfach verkauft
- Verkaufsjahr fehlt
- Verkaufsjahr außerhalb Raster
- Verkaufspreis fehlt
- Aufteilung fehlt: weder Anteil G+B noch Verkehrswertanteil
- § 6b unzulässig: § 6b ja, aber Vorbesitzzeit < par\_6bVorbesitz

Weil der Verkauf zum 31.12. gilt, reicht Verkaufsjahr − Kaufjahr ≥ 6 für die sechs Jahre Vorbesitz.

**Verknüpfung zur Prognose:** Das aktiv-Flag liest das Verkaufsjahr per INDEX/MATCH aus vk\_Jahr; prg\_Bestand ist schon im Verkaufsjahr 0. Der Veräußerungsgewinn Q fließt in den Rücklagenspiegel (Abschnitt 12): bei § 6b ja in die Rücklage, sonst als sofort steuerwirksam. Noch nicht umgesetzt ist ein Verkaufspreis als Faktor × Jahresmiete.

**Abnahme Etappe 4 (geprüft in `pruefung/pruefen.py`):**
- Preis 1,4 Mio, Kosten 0, Buchwert gesamt 680.000 (Gebäude 480.000 Ende 2027, G+B 200.000), Verkehrswertanteil Gebäude 50 %:
  - Veräußerungsgewinn 720.000
  - davon Gebäude 220.000 und G+B 500.000
- Kaufvertrag 30 % G+B, Kosten 40.000, Verkauf Ende 2030:
  - Gewinn Gebäude 536.000
  - Gewinn G+B 208.000

## 12. Rücklagenspiegel § 6b (Etappe 5)

Bei 6b-Nutzung ja wird der Veräußerungsgewinn in eine Rücklage eingestellt, getrennt nach Gebäude und Grund und Boden, weil die Übertragbarkeit unterschiedlich ist. Die Rücklage stundet die Steuer bis zur Reinvestition oder bis zum Fristablauf.

Das Blatt Rücklagen hat zwei Teile: links eine Zeile je Verkauf, rechts den Spiegel je Jahr. Zeile n links gehört zu Zeile n im Blatt Verkäufe und ist nur gefüllt, wenn dort § 6b ja steht und der Status OK lautet.

**Teil 1: je Rücklage (umgesetzt)**

| Spalte | Feld | Name | Formel |
| --- | --- | --- | --- |
| A | RücklageID | rl\_ID | "RL-" & ObjektID |
| B | ObjektID Herkunft | rl\_ObjektID | aus Verkäufe |
| C | Bildungsjahr | rl\_Jahr | Verkaufsjahr, Bildung zum 31.12. |
| D | Rücklage Gebäude | rl\_Geb | MAX(Gewinn Gebäude; 0) |
| E | Rücklage G+B | rl\_GuB | MAX(Gewinn G+B; 0) |
| F | Rücklage gesamt | rl\_Betrag | D + E |
| G | Fristjahr | rl\_Fristjahr | C + par\_6bFrist, bei Neubau begonnen + par\_6bFristNeubau |
| H | übertragen Gebäude | rl\_UebGeb | Summe ü1 aus Neuobjekte (Abschnitt 13) |
| I | übertragen G+B | rl\_UebGuB | Summe ü2 + ü3 aus Neuobjekte |
| J | Auflösung im Fristjahr | rl\_Aufloesung | F − H − I |
| K | Gewinnzuschlag | rl\_Zuschlag | J × par\_6bZuschlag × (G − C) |
| L | Hinweis | rl\_Hinweis | "Frist endet nach Prognoseende", wenn G > par\_Endjahr und J > 0 |

```text
D2  =IF(AND(INDEX(vk_Status,1)="OK", INDEX(vk_6b,1)="ja"), MAX(INDEX(vk_GewinnGeb,1),0), "")
G2  =IF(…, C2 + IF(INDEX(vk_6bNeubau,1)="ja", par_6bFristNeubau, par_6bFrist), "")
K2  =IF(…, J2 * par_6bZuschlag * (G2 - C2), "")
```

**Nur positive Teilgewinne.** Gebäude und G+B sind getrennte Wirtschaftsgüter. Ein Verlust beim Gebäude mindert die Rücklage aus dem G+B-Gewinn nicht, er wirkt im Verkaufsjahr sofort.

**Frist.** Regelfrist vier Jahre, sechs Jahre, wenn mit dem Bau eines neuen Gebäudes vor Ende des vierten Jahres begonnen wurde (§ 6b Abs. 3, geklärt). Das steuert die Eingabe „§ 6b Neubau begonnen“ im Blatt Verkäufe. Ohne Reinvestition wird die Rücklage am Ende des Fristjahrs aufgelöst, mit 6 % Zuschlag je vollem Jahr ihres Bestehens: 24 % bei vier, 36 % bei sechs Jahren.

**Teil 2: Spiegel je Jahr (umgesetzt)**, Prognosejahre 2027 bis 2046:

| Feld | Name | Formel |
| --- | --- | --- |
| Jahr | rls\_Jahr | par\_Startjahr fortlaufend |
| Veräußerungsgewinne | rls\_Gewinne | SUMIFS(vk\_Gewinn) im Jahr, Status OK oder § 6b unzulässig |
| Einstellung in Rücklage | rls\_Bildung | SUMIFS(rl\_Betrag, rl\_Jahr) |
| Übertragung auf Neuobjekte | rls\_Uebertragung | SUMIFS(ne\_UeGesamt, ne\_Kaufjahr) |
| Auflösung | rls\_Aufloesung | SUMIFS(rl\_Aufloesung, rl\_Fristjahr) |
| Gewinnzuschlag | rls\_Zuschlag | SUMIFS(rl\_Zuschlag, rl\_Fristjahr) |
| steuerwirksam aus Verkauf und Rücklage | rls\_Steuerwirksam | Gewinne − Einstellung + Auflösung + Zuschlag |
| Rücklagenbestand Ende | rls\_Bestand | Vorjahr + Einstellung − Übertragung − Auflösung |

rls\_Steuerwirksam ist die Brücke zur Steuer in Etappe 7: Steuer aus Verkäufen = rls\_Steuerwirksam × par\_Steuersatz. Bei „§ 6b unzulässig“ ist der Gewinn korrekt berechnet, geht aber ohne Rücklage sofort in die Steuer. Verkaufszeilen mit anderen Fehlern zählen nicht.

**Abnahme Etappe 5 (geprüft in `pruefung/pruefen.py`):**
- Gewinn 720.000 aus dem Abnahmefall der Etappe 4, § 6b ja, Verkauf Ende 2027:
  - Rücklage 720.000, davon Gebäude 220.000 und G+B 500.000
  - steuerwirksam 2027 null, Bestand Ende 2027 bis 2030 720.000
  - Fristjahr 2031: Auflösung 720.000, Zuschlag 172.800, steuerwirksam 892.800, Bestand 0
- Neubau begonnen: Fristjahr 2033, Zuschlag 259.200
- § 6b nein: Gewinn 744.000 sofort steuerwirksam, keine Rücklage
- Kaufvertrag 80 % G+B: Gebäude −136.000, G+B 920.000; Rücklage 920.000, steuerwirksam −136.000
- § 6b mit Vorbesitz 5 Jahre: keine Rücklage, Gewinn 368.000 sofort steuerwirksam
- Verkauf 2044: Fristjahr 2048 mit Hinweis, Bestand Ende 2046 1.008.000

**Offen:**
- Bestehende Rücklagen aus Verkäufen vor 2027 (Bestand zum Ende des Basisjahrs) fehlen noch. Der Spiegel startet mit 0.
- Eine freiwillige Auflösung vor dem Fristjahr ist nicht vorgesehen.

## 13. Reinvestition (Etappe 6)

Ein Neuobjekt nimmt die Rücklage auf. Der übertragene Gewinn mindert die AfA-Basis des neuen Gebäudes, dadurch läuft die AfA künftig von einem niedrigeren Wert. Das ist der Kern der Stundung: keine Steuer heute, dafür weniger Abschreibung morgen.

**Spaltenlayout Blatt Neuobjekte (umgesetzt)**

| Spalte | Feld | Name | Quelle |
| --- | --- | --- | --- |
| A | NeuID | ne\_ID | Eingabe, eindeutig, nicht gleich einer ObjektID |
| B | Name | ne\_Name | Eingabe, optional |
| C | Kaufjahr | ne\_Kaufjahr | Eingabe, Kauf zum Jahresende |
| D | Kaufpreis | ne\_Kaufpreis | Eingabe |
| E | Anteil G+B | ne\_AnteilGuB | Eingabe, Prozent |
| F | Kaufnebenkosten | ne\_Nebenkosten | Eingabe, leer = 0 |
| G | AfA-Satz | ne\_AfASatz | Eingabe; bei degressiver AfA bestimmt er die Nutzungsdauer (1 / Satz) |
| H | AfA-Methode | ne\_AfAMethode | Dropdown linear oder degressiv, leer = linear |
| I | Mietrendite auf Kaufpreis | ne\_Mietrendite | Eingabe, optional |
| J | Erhaltung auf Kaufpreis | ne\_ErhQuote | Eingabe, optional |
| K | Quelle RücklageID | ne\_Quelle | Dropdown aus rl\_ID, optional |
| L | im Modell | ne\_Gueltig | 1 bei vollständigen Pflichtfeldern, eindeutiger ID, Kaufjahr im Raster |
| M | AK G+B neu | ne\_AKGuBNeu | (D + F) × E |
| N | AK Gebäude neu | ne\_AKGebNeu | (D + F) × (1 − E) |
| O | Rücklage Gebäude verfügbar | ne\_RLGeb | rl\_Geb minus ü1 der Zeilen darüber mit gleicher Quelle |
| P | Rücklage G+B verfügbar | ne\_RLGuB | rl\_GuB minus ü2 und ü3 der Zeilen darüber |
| Q | ü1 | ne\_Ue1 | MIN(O; N) |
| R | ü2 | ne\_Ue2 | MIN(P; M) |
| S | ü3 | ne\_Ue3 | MIN(P − R; N − Q) |
| T | übertragen gesamt | ne\_UeGesamt | Q + R + S |
| U | AfA-Basis Gebäude | ne\_AfABasis | N − Q − S |
| V | steuerliche AK G+B | ne\_AKGuB | M − R |
| W | Status | ne\_Status | Plausibilität |

Nebenkosten wie Grunderwerbsteuer und Notar werden aktiviert und im Verhältnis des Kaufpreises auf G+B und Gebäude verteilt. Die verfügbare Rücklage (O, P) ist nur gefüllt, wenn das Kaufjahr zwischen Bildungsjahr und Fristjahr der Quelle liegt, sonst 0. Die AfA-Methode ist linear (Satz × AfA-Basis) oder degressiv nach § 7 Abs. 5a EStG.

**Übertragung in fester Reihenfolge (korrigiert).** Der erste Entwurf übertrug nur den Gebäudegewinn und nur auf das Gebäude (`I2 = MIN(Rücklage Spalte E; H2)`). Der G+B-Gewinn ging dabei verloren. Richtig nach § 6b Abs. 1 EStG ist: Ein Gebäudegewinn darf nur auf ein Gebäude übertragen werden, ein G+B-Gewinn auf Gebäude oder auf G+B. Steuerlich günstig ist diese Reihenfolge:

1. **ü1 = MIN(Rücklage Gebäude; AK Gebäude neu).** Der Gebäudegewinn geht auf das neue Gebäude, denn er hat keine andere Verwendung.
2. **ü2 = MIN(Rücklage G+B; AK G+B neu).** Der G+B-Gewinn geht zuerst auf den neuen G+B. Das kostet keine AfA, die Steuer bleibt bis zum Verkauf des Grundstücks gestundet.
3. **ü3 = MIN(Rücklage G+B − ü2; AK Gebäude neu − ü1).** Erst der Rest des G+B-Gewinns geht auf das Gebäude und mindert dessen AfA-Basis.

Danach gilt:
- AfA-Basis Gebäude = AK Gebäude neu − ü1 − ü3
- steuerliche AK G+B = AK G+B neu − ü2, Kürzung bis auf 0 ist zulässig (geklärt)
- Rücklage Gebäude sinkt um ü1, Rücklage G+B um ü2 + ü3. Was übrig bleibt, wartet auf ein weiteres Neuobjekt oder wird im Fristjahr mit Zuschlag aufgelöst.

Beispiel: G+B-Gewinn 4, neuer G+B kostet 3. Dann werden 3 beim G+B abgezogen und 1 beim Gebäude.

**Mehrere Neuobjekte je Rücklage.** Je Neuobjekt gibt es eine Quelle. Nutzen mehrere Neuobjekte dieselbe Rücklage, bedienen sie sich in Zeilenreihenfolge: Jede Zeile sieht nur, was die Zeilen darüber übrig gelassen haben. Die Zeilen sollten daher chronologisch stehen.

**Status je Zeile**, in dieser Reihenfolge:
- Pflichtfeld fehlt (NeuID, Kaufjahr, Kaufpreis, Anteil G+B, AfA-Satz)
- NeuID doppelt
- NeuID wie Bestandsobjekt
- Kaufjahr außerhalb Raster
- Rücklage unbekannt, keine Übertragung
- Kauf vor Bildung der Rücklage, keine Übertragung
- Kauf nach Fristjahr, keine Übertragung

Die ersten vier nehmen das Objekt aus dem Modell (ne\_Gueltig = 0). Bei den letzten drei bleibt es im Modell, nur ohne Übertragung.

**Prognose.** Die neue AfA-Basis fließt zurück in die Objektlogik: Jedes Neuobjekt hat in der Prognosematrix einen eigenen Block mit 20 Jahreszeilen nach den 200 Objektblöcken, prg\_Neu = 1. Kauf zum Jahresende heißt:
- im Kaufjahr: Buchwert = AfA-Basis, im Bestand, noch keine Miete, Erhaltung und AfA
- ab dem Folgejahr: AfA = MIN(AfA-Basis × Satz; Vorjahresbuchwert); Miete = Kaufpreis × Mietrendite × (1 + Mietsteigerung)^(Jahr − Kaufjahr), Erhaltung entsprechend
- Verkehrswert = Kaufpreis, ab dem Kaufjahr mit der Wertsteigerung fortgeschrieben
- degressive AfA (Methode „degressiv“): AfA = MIN(MAX(Vorjahresbuchwert × par\_AfADegressiv; Vorjahresbuchwert / Restnutzungsdauer); Vorjahresbuchwert). Die Restnutzungsdauer zu Jahresbeginn ist 1 / AfA-Satz − (Jahr − Kaufjahr − 1), mindestens 1. Das bildet den Wechsel zur linearen AfA nach § 7 Abs. 5a Satz 4 EStG ab, sobald dieser günstiger ist; danach bleibt die AfA gleich. Die AfA-Basis nach Übertragung der Rücklage ist dieselbe wie bei linearer AfA.

Ein Neuobjekt gehört nur in die Plan-Linie der Übersicht, nicht in die Baseline. Die Baseline summiert deshalb nur Zeilen mit prg\_Neu = 0.

**Rückkopplung in den Rücklagenspiegel (umgesetzt):** Das Rücklagenblatt hat die Spalten „übertragen Gebäude“ rl\_UebGeb (Summe ü1) und „übertragen G+B“ rl\_UebGuB (Summe ü2 + ü3) je Quelle. Die Auflösung im Fristjahr ist die Rücklage minus das Übertragene, der Zuschlag rechnet nur auf diesen Rest. Der Spiegel je Jahr hat die Spalte rls\_Uebertragung (ne\_UeGesamt im Kaufjahr), die den Bestand mindert. Übertragung ist nicht steuerwirksam. Ist die Rücklage voll übertragen, entfallen Auflösung und Zuschlag.

**Abnahme Etappe 6:**
- Ausgangslage: Rücklage aus dem Abnahmefall der Etappe 4 (Gebäude 220.000, G+B 500.000). Neuobjekt mit Kaufpreis 1,2 Mio, Anteil G+B 30 %, also G+B 360.000 und Gebäude 840.000.
- Übertragung: ü1 = 220.000, ü2 = 360.000, ü3 = 140.000.
- Ergebnis: AfA-Basis Gebäude 480.000, steuerliche AK G+B 0, Restrücklage 0.
- Die AfA des Neuobjekts läuft von 480.000, nicht von 840.000.

Weitere geprüfte Fälle in `pruefung/pruefen.py`:
- Kauf Ende 2028, AfA 3 %: Buchwert 2028 480.000, AfA 2029 14.400, Miete 2029 61.200; Plan 2028 1,2 Mio, Baseline ohne Neuobjekt
- zwei Neuobjekte teilen sich die Rücklage: das erste nimmt 500.000, das zweite (Nebenkosten 40.000, G+B 260.000) den Rest 220.000 als ü2; AfA-Basis 780.000, AK G+B 40.000
- Teilübertragung 300.000: Rest 420.000 wird 2031 mit Zuschlag 100.800 aufgelöst
- Statusfälle, darunter Kauf nach Fristjahr und vor Bildung der Rücklage
- degressive AfA 5 %, AfA-Basis 800.000, Nutzungsdauer 33⅓ Jahre: 2028 40.000, 2029 38.000, 2041 20.533,68; ab 2042 linear 20.179,65 (Restnutzungsdauer 19⅓ < 20), Buchwert 2046 289.241,71. Linear 3 % zum Vergleich 24.000, Buchwert 2046 344.000. Bei Nutzungsdauer 20 Jahre ist linear ab dem zweiten Jahr höher, die AfA bleibt 40.000.

**Offen:**
- Verkauf eines Neuobjekts innerhalb des Rasters
- mehrere Quellen für ein Neuobjekt
- Übertragung auf Anschaffungen im Vorjahr der Veräußerung (§ 6b Abs. 1)
- „§ 6b Neubau begonnen“ aus dem Neuobjekt ableiten statt im Blatt Verkäufe eingeben

## 14. Liquidität und Auswertung (Etappe 7)

Die Prognosematrix liefert je Objekt und Jahr die Einzelwerte. Liquidität und Auswertung fassen sie über alle Objekte zu Jahreswerten zusammen, per SUMMEWENNS über das Jahr. Seit Etappe 8 haben beide Blätter je Szenario eine Tabelle mit denselben Spalten, nebeneinander mit einer Spalte Abstand: A Plan, B, C, Baseline (Abschnitt 15). Zeile 1 trägt den Tabellentitel, Zeile 2 die Kopfzeile, ab Zeile 3 je Prognosejahr eine Zeile. Alle Angaben sind vor Finanzierung.

**Blatt Liquidität, Szenario A (umgesetzt)**

| Spalte | Feld | Name | Formel |
| --- | --- | --- | --- |
| A | Jahr | liq\_Jahr | 2027 bis 2046 |
| B | Mieten und weitere Einnahmen | liq\_Einnahmen | Summe prg\_Miete + prg\_Einnahmen |
| C | Erhaltung und weitere Ausgaben | liq\_Ausgaben | Summe prg\_Erhaltung + prg\_Ausgaben |
| D | AfA Gebäude | liq\_AfA | Summe prg\_AfA |
| E | laufendes Ergebnis | liq\_Ergebnis | B − C − D, gleich der Summe prg\_Ergebnis |
| F | steuerwirksam aus Verkauf und Rücklage | liq\_Verkauf | rls\_Steuerwirksam des Jahres |
| G | Zinsertrag Alternativanlage | liq\_Zins | Liquidität Vorjahr × par\_Alternativrendite (Etappe 8) |
| H | Ergebnis vor Verlustvortrag | liq\_ZvE | E + F + G |
| I | Verlustvortrag genutzt | liq\_VortragGenutzt | MIN(Vortrag Vorjahr; MAX(H; 0)) |
| J | Bemessungsgrundlage | liq\_Bemessung | MAX(H; 0) − I |
| K | Verlustvortrag Ende | liq\_Vortrag | Vortrag Vorjahr − I + MAX(−H; 0) |
| L | Steuer | liq\_Steuer | J × par\_Steuersatz |
| M | Verkaufserlöse netto | liq\_Verkaufserloes | vk\_Nettoerloes der Verkäufe mit Status OK oder „§ 6b unzulässig“ |
| N | davon Buchwert-Rückfluss | liq\_Rueckfluss | M − rls\_Gewinne, also Buchwert Gebäude + AK G+B |
| O | Kauf Neuobjekte inkl. Nebenkosten | liq\_Kauf | Kaufpreis + Nebenkosten der gültigen Neuobjekte im Kaufjahr |
| P | freier Mittelzufluss | liq\_Zufluss | B − C + G + M − L − O |
| Q | Liquidität kumuliert Ende | liq\_Kum | Vorjahr + P |

**Blatt Liquidität, Baseline (umgesetzt):** gleiche Spalten, Namen lqb\_…. Einnahmen und Ausgaben sind die Basiswerte aller Objekte mit ihrer Steigerungsrate, unabhängig von Verkäufen. Die AfA ist der Rückgang von prg\_BuchwertHalten gegenüber dem Vorjahr, im ersten Jahr gegenüber der Summe der Restbuchwerte. Verkauf, Erlöse und Kauf sind 0.

**Zur Steuer**

Die Steuer hängt am 6b-Schalter des jeweiligen Verkaufs. Bei Rücklage ist sie im Verkaufsjahr null, der Gewinn ist gestundet; Auflösung und Zuschlag erhöhen sie im Fristjahr. Ohne Rücklage fällt sie sofort an. All das fasst der Rücklagenspiegel in rls\_Steuerwirksam zusammen (Abschnitt 12).

Ein Verlust, etwa aus dem Gebäudeteil eines Verkaufs, ergibt keine negative Steuer. Er wird vorgetragen und mit den nächsten Gewinnen verrechnet; das entspricht der GmbH, deren Verluste nur mit eigenen Gewinnen verrechnet werden. Die Mindestbesteuerung (§ 10d Abs. 2 EStG, § 10a GewStG: über 1 Mio nur zu 60 %) und der Verlustrücktrag fehlen noch.

Seit Etappe 8 wird die Liquidität mit par\_Alternativrendite verzinst (Abschnitt 15).

**Blatt Auswertung, Szenario A (umgesetzt)**

| Spalte | Feld | Name | Formel |
| --- | --- | --- | --- |
| A | Jahr | aus\_Jahr | 2027 bis 2046 |
| B | laufendes Ergebnis | aus\_Ergebnis | aus Liquidität |
| C | steuerwirksam aus Verkauf und Rücklage | aus\_Verkauf | aus Liquidität |
| D | Zinsertrag Alternativanlage | aus\_Zins | aus Liquidität |
| E | Gesamt-GuV vor Steuern | aus\_GuV | B + C + D |
| F | Steuer | aus\_Steuer | aus Liquidität |
| G | Ergebnis nach Steuern | aus\_NachSteuer | E − F |
| H | Steuer kumuliert | aus\_SteuerKum | laufende Summe über F |
| I | Verkehrswert Bestand | aus\_Verkehrswert | prg\_Verkehrswert der Zeilen mit prg\_Bestand = 1 |
| J | Buchwert Bestand | aus\_Buchwert | prg\_Buchwert + prg\_BuchwertGuB, ebenso gefiltert |
| K | stille Reserven | aus\_StilleReserven | I − J |
| L | § 6b-Rücklage Bestand | aus\_Ruecklage | rls\_Bestand |
| M | Verlustvortrag | aus\_Vortrag | aus Liquidität |
| N | Liquidität kumuliert | aus\_Liquiditaet | aus Liquidität |
| O | Gesamtvermögen vor latenter Steuer | aus\_Vermoegen | I + N |
| P | latente Steuer | aus\_LatenteSteuer | MAX(K + L − M; 0) × par\_Steuersatz |
| Q | Gesamtvermögen nach latenter Steuer | aus\_VermoegenNetto | O − P |

**Blatt Auswertung, Baseline (umgesetzt):** gleiche Spalten, Namen asb\_…. Gezählt werden alle Bestandsobjekte (prg\_Neu = 0) unabhängig vom Verkauf, der Gebäudebuchwert aus prg\_BuchwertHalten, Rücklage 0.

**Stille Reserven und latente Steuer**

Stille Reserven zeigen, wie viel unversteuerter Wert im Bestand steckt: Verkehrswert minus Buchwert über alle Objekte, die am Jahresende noch im Bestand sind. Beim Neuobjekt stecken die übertragenen Gewinne darin, weil der Buchwert um sie gemindert ist. Die nicht übertragene Rücklage ist ebenfalls gestundete Steuer. Die latente Steuer ist deshalb die Steuer auf stille Reserven plus Rücklage, gemindert um den Verlustvortrag. Sie gilt für einen gedachten Verkauf aller Objekte zum Verkehrswert ohne neue Rücklage und ohne Gewinnzuschlag.

Erst das Gesamtvermögen nach latenter Steuer macht Halten und Verkaufen vergleichbar: Die Baseline hat höhere stille Reserven, der Plan hat Liquidität, aber schon Steuer gezahlt.

Die Zahlen der Abnahme Etappe 7 gelten ohne Zins auf die Liquidität; die Prüffälle setzen dafür par\_Alternativrendite = 0.

**Übersicht (umgesetzt):** Neben dem Verkehrswert (Spalten B bis D) zeigt die Tabelle das Gesamtvermögen nach latenter Steuer für Baseline und Plan und die Differenz (E bis G, ueb\_VermBaseline, ueb\_VermPlan, ueb\_VermDifferenz), mit einem zweiten Diagramm. Im Basisjahr ist es für beide gleich: Verkehrswert minus latente Steuer auf Verkehrswert − Restbuchwert − AK G+B.

**Abnahme Etappe 7:**
- Zwei Objekte ohne Verkauf, 2027: laufendes Ergebnis 72.980 = 37.000 + 35.980, gleich der Einzelsumme aus der Prognose. Plan und Baseline liefern dieselbe Steuer (21.894), dieselbe Liquidität und dasselbe Gesamtvermögen.
- Verkauf Ende 2030 ohne § 6b, Gewinn 744.000: Steuer 2030 = (40.115,43 + 744.000) × 30 % = 235.234,63, also 223.200 mehr als ohne Verkauf. Der Buchwert-Rückfluss ist 616.000 = 416.000 + 200.000.

Weitere geprüfte Fälle in `pruefung/pruefen.py`:
- Rücklage ohne Reinvestition: Steuer 2027 nur auf das laufende Ergebnis, 2031 Steuer 267.840 auf Auflösung und Zuschlag; bis dahin steht die Rücklage in der latenten Steuer
- Reinvestition Ende 2028: Kauf 1,2 Mio mindert die Liquidität, das Gesamtvermögen nach latenter Steuer bleibt gleich (1.225.900), die gestundete Steuer wandert aus der Rücklage in die stillen Reserven des Neuobjekts
- Verlustvortrag: Verlust 2030 von 95.884,57 wird 2034 mit der Auflösung verrechnet

**Offen:**
- Mindestbesteuerung und Verlustrücktrag
- Grunderwerbsteuer und Nebenkosten des Neuobjekts nur über die Eingabe Kaufnebenkosten; par\_GrESt wird noch nicht verwendet
- Ein Verkauf mit anderem Status als OK oder „§ 6b unzulässig“ nimmt das Objekt aus dem Plan, bringt aber keinen Erlös. Der Status ist rot, bis die Eingabe vollständig ist.

## 15. Szenariovergleich (Etappe 8)

Der Vergleich beantwortet die Kernfrage des Mandanten: Lohnt die § 6b-Kette, oder ist es besser, die Steuer sofort zu zahlen und das freie Kapital anderweitig anzulegen? Alle Pfade laufen über dieselben Objekte, Verkäufe und 20 Jahre, nur die Behandlung des Veräußerungsgewinns und der Neuobjekte unterscheidet sich.

**Die Szenarien**

| | Verkauf mit § 6b ja | Neuobjekt mit Quelle-Rücklage | Neuobjekt ohne Quelle | freie Mittel |
| --- | --- | --- | --- | --- |
| A Plan, § 6b-Kette | Rücklage wie in Abschnitt 12 | Kauf mit Übertragung, geminderte AfA-Basis | Kauf | Alternativanlage |
| B sofort versteuern, anlegen | Gewinn sofort steuerpflichtig | entfällt, das Geld bleibt angelegt | Kauf | Alternativanlage |
| C sofort versteuern, kaufen | Gewinn sofort steuerpflichtig | Kauf ohne Übertragung, volle AfA-Basis | Kauf | Alternativanlage |
| Baseline alles halten | kein Verkauf | kein Kauf | kein Kauf | Alternativanlage |

A − C zeigt die reine Wirkung von § 6b, A − B die Frage Immobilie oder Geldanlage, A − Baseline die Frage Umschichten oder Halten.

**Umsetzung: alle Szenarien gleichzeitig (umgesetzt)**

Statt eines Schalters mit gespeicherten Läufen rechnen alle Szenarien zugleich in Formeln. So ist der Vergleich immer aktuell und braucht kein Makro. Der Schalter par\_Szenario entfällt. Möglich ist das, weil sich B und C von A nur an zwei Stellen unterscheiden:

- **Steuer auf den Verkauf:** A nimmt rls\_Steuerwirksam aus dem Rücklagenspiegel, B und C nehmen rls\_Gewinne, also jeden Veräußerungsgewinn im Verkaufsjahr.
- **Neuobjekte:** Die Prognose hat vier Zusatzspalten. prg\_MitQuelle ist 1 für Neuobjekte mit Quelle-Rücklage (ne\_MitQuelle); B summiert nur Zeilen mit 0. prg\_AfAOhne6b, prg\_BuchwertOhne6b und prg\_BuchwertGuBOhne6b rechnen das Neuobjekt mit den vollen AK (ne\_AKGebNeu, ne\_AKGuBNeu), nach derselben AfA-Methode. Für Bestandsobjekte sind sie gleich prg\_AfA, prg\_Buchwert und prg\_BuchwertGuB.

Liquidität und Auswertung haben je Szenario eine Tabelle mit denselben Spalten wie in Abschnitt 14:

| Szenario | Liquidität | Auswertung | Zeilen der Prognose | AfA und Buchwert | Verkauf | Kauf |
| --- | --- | --- | --- | --- | --- | --- |
| A | liq\_… | aus\_… | alle | prg\_AfA, prg\_Buchwert, prg\_BuchwertGuB | rls\_Steuerwirksam | alle gültigen Neuobjekte |
| B | lvb\_… | avb\_… | prg\_MitQuelle = 0 | …Ohne6b | rls\_Gewinne | nur ne\_MitQuelle = 0 |
| C | lvc\_… | avc\_… | alle | …Ohne6b | rls\_Gewinne | alle gültigen Neuobjekte |
| Baseline | lqb\_… | asb\_… | Bestandsobjekte | prg\_BuchwertHalten | 0 | 0 |

Die Rücklage steht nur in A in der latenten Steuer. B und C haben keine.

**Alternativanlage**

- In allen Szenarien liegen die freien Mittel in einer Alternativanlage mit par\_Alternativrendite (Standard 3 %, wie in der Planungsreferenz). Auch in A und C, damit die Differenz nur die Wirkung der Kette zeigt.
- Zins = Liquidität am Vorjahresende × Rendite. Der Mittelzufluss gilt zum Jahresende, der Zins also ab dem Folgejahr.
- Der Zinsertrag ist voll steuerpflichtig und geht in das Ergebnis vor Verlustvortrag ein.
- Negative Liquidität kostet denselben Satz, vor Finanzierung gibt es keinen eigenen Kreditzins.

**Blatt Vergleich (umgesetzt)**

Das Blatt steht direkt hinter der Übersicht.

| Spalte | Inhalt | Name |
| --- | --- | --- |
| A | Kennzahl | vg\_Kennzahl |
| B–E | A Plan, B, C, Baseline | vg\_A, vg\_B, vg\_C, vg\_Baseline |
| F–H | A − B, A − C, A − Baseline | vg\_DiffB, vg\_DiffC, vg\_DiffBaseline |
| I | Erläuterung | |

Die Kennzahlen in dieser Reihenfolge (Zeile im Bereich in Klammern):

- **Bestände am Ende des letzten Prognosejahrs:**
  - Endvermögen nach latenter Steuer (0)
  - Verkehrswert (1)
  - Liquidität (2)
  - latente Steuer (3)
  - Buchwert (4)
  - stille Reserven (5)
  - Rücklage (6)
  - Verlustvortrag (7)
- **Summen über alle Jahre:**
  - Einnahmen (8)
  - Ausgaben (9)
  - AfA (10)
  - laufendes Ergebnis (11)
  - steuerwirksam aus Verkauf und Rücklage (12)
  - Zinsertrag (13)
  - Steuer (14)
  - Verkaufserlöse (15)
  - Kauf Neuobjekte (16)

Darunter steht eine Tabelle mit dem Endvermögen nach latenter Steuer je Jahr und Szenario (vgj\_Jahr, vgj\_A, vgj\_B, vgj\_C, vgj\_Baseline) mit Liniendiagramm.

**Die ehrliche Kennzahl**

Entscheidend ist das Endvermögen nach latenter Steuer, nicht die gesparte Steuer allein. Szenario A spart Steuer heute, verliert aber AfA und bindet Kapital in Immobilien. Ohne den Abzug der latenten Steuer wäre A geschönt, weil die gestundete Steuer nie auftauchte.

**Abnahme Etappe 8:**
- Jeder Pfad liefert ein Endvermögen. Die Differenz ist nachvollziehbar aus gestundeter Steuer, verlorener AfA und Alternativrendite.
- Bei Alternativrendite null liegen A und C gleichauf: § 6b spart keine Steuer, sondern stundet sie, und ohne Zins ist die Stundung nichts wert. Das gilt, solange kein Verlustvortrag verfällt und die latente Steuer nicht bei 0 gekappt wird.
- Mit positiver Alternativrendite liegt A vor C, der Vorsprung ist der Zins auf die gestundete Steuer.

*Korrektur gegenüber dem ersten Entwurf:* Dort stand, bei Alternativrendite null müsse A vorn liegen. Das gilt nur gegenüber B, weil das Neuobjekt Miete und Wertsteigerung bringt. Gegenüber C ist es ein Gleichstand.

Im Prüfskript, Abnahmefall der Reinvestition (Verkauf Ende 2027, Gewinn 720.000, Neuobjekt Ende 2028 für 1,2 Mio). Die Sollwerte stammen aus einem eigenen Nachbau außerhalb der Mappe:

| | A | B | C |
| --- | --- | --- | --- |
| Steuer 2027 | 11.100 | 227.100 | 227.100 |
| AfA Neuobjekt ab 2029 | 14.400 | – | 25.200 |
| Endvermögen 2046, Rendite 0 % | 2.310.183,85 | 1.225.900 | 2.310.183,85 |
| Endvermögen 2046, Rendite 3 % | 2.615.589,88 | 1.819.466,84 | 2.522.678,66 |

- Bei 3 % ist A − C = 92.911,22. Der Zins 2028 in A ist 1.441.900 × 3 % = 43.257.
- Ohne § 6b-Rücklage sind A, B und C gleich, ohne Verkauf ist A gleich der Baseline, auch mit Zins.
- Ein Neuobjekt ohne Quelle-Rücklage bleibt in B erhalten.

**Offen:**
- Steuersatz der Alternativanlage: Kapitalerträge der GmbH sind voll steuerpflichtig; eine Anlage in Aktien (§ 8b KStG) wäre günstiger und ist nicht abgebildet.
- ~~Varianten der Eingaben speichern und vergleichen~~: umgesetzt in Etappe 9 als Blatt Varianten mit dem Makro „Variante festhalten“ (Abschnitt 19).

## 16. Datenanbindung: zwei Quellen

Das Modell speist sich aus zwei getrennten Quellen, weil die laufende Buchhaltung nicht alle steuerlichen Stammdaten enthält. Beide fließen beim Generieren ins Objektblatt, in zwei klar getrennte Bereiche.

**Quelle 1: laufende Werte (automatisch)**

- Herkunft: die vorhandene Kanzlei-Excel mit einem Blatt je Kostenstelle, also je Immobilie.
- Inhalt: laufende Buchwerte, Einnahmen und Ausgaben. Keine steuerlichen Stammdaten.
- Einlesen: das Python-Skript geht alle Blätter durch und liest je Blatt die immer gleichen Zellen aus. Jedes Blatt wird eine Objektzeile.
- Voraussetzung: einheitliches Blattlayout und eine eindeutige ObjektID je Blatt an fester Stelle.

**Quelle 2: steuerliche Stammdaten (gepflegt)**

- Inhalt: Anschaffungskosten getrennt nach Gebäude und Grund und Boden, AfA-Satz, Kaufjahr, Restbuchwert als Startwert.
- Herkunft: idealerweise das Anlageverzeichnis; sonst einmalig von Hand gepflegt, da selten änderlich.
- Grund: ohne diese Werte sind weder AfA-Fortschreibung noch § 6b-Aufteilung möglich.

**Layout eines Kostenstellenblatts (Muster liegt vor)**

Die Zeilen folgen der DATEV-BWA Form 01 (Kurzfristige Erfolgsrechnung). Die Nummern in Spalte B sind BWA-Zeilennummern, keine Sachkonten. Die BWA ist kontenrahmenunabhängig; ob dahinter SKR03 oder SKR04 gebucht wird, geht aus dem Blatt nicht hervor.

| Zelle/Bereich | Inhalt |
| --- | --- |
| B2 | Kostenstelle, z. B. „KSt 1“ |
| C2 | Objektbezeichnung, z. B. „KC 24+26“ |
| Zeile 4 | Kopf: B „Nr.“, C „Bezeichnung kurz“, danach Jahres- und Monatsspalten (siehe unten) |
| ab Zeile 6 | eine Zeile je BWA-Position, Schlüssel ist die Nummer in Spalte B |
| G–R | Ist-Werte bis zum letzten gebuchten Monat, danach Hochrechnung per Mittelwert |

Zwei Kopfvarianten kommen vor, die Einleseschicht liest beide:

| Variante | Spalten in Zeile 4 |
| --- | --- |
| erstes Muster | F 2025, G–R Monate 2026, S 2026, T–AM 2027–2046 (Jahre als Zahl) |
| Planungsreferenz (Zielbild, Abschnitt 18) | F „Jahr 2024“, G „Jahr 2025“, H–S Monate 2026 als Datum, T „Jahr 2026“, U „Plan 2027“, V–AN 2028–2046 |

Eine Spalte gilt als Jahresspalte, wenn der Kopf eine Jahreszahl ist oder „Jahr 2026“ bzw. „Plan 2027“ lautet. Monatsspalten (Datum) werden nie als Jahr gelesen. Ein Summenblatt über alle Kostenstellen (B2 nur „KSt“ oder C2 „Alle Objekte“) wird übersprungen, damit es nicht als eigenes Objekt zählt.

Für das Modell relevante BWA-Zeilen (über die Nummer in Spalte B suchen, nicht über die Zeilennummer):

| BWA-Nr. | Bezeichnung | Verwendung im Modell |
| --- | --- | --- |
| 1020 | Umsatzerlöse | Miete im Basisjahr |
| 1090 | So. betr. Erlöse | weitere laufende Einnahmen |
| 1100–1220, 1260 | Personal, Raum, betr. Steuern, Versicherungen, Besondere Kosten (1160), Kfz, Werbung, Warenabgabe, Sonstige | weitere laufende Ausgaben |
| 1240 | Abschreibungen | Abgleich mit der AfA-Fortschreibung, nicht als Eingabe |
| 1250 | Reparatur/Instandh. | Erhaltungsaufwand im Basisjahr |
| 1310, 1322 | Zinsaufwand, Zinserträge | erst Stufe 2 (Finanzierung) |
| 1300, 1345, 1380 | Betriebsergebnis, Ergebnis vor Steuern, Vorläufiges Ergebnis | Plausibilisierung |

Auffälligkeiten im Muster, vor dem Einlesen mit der Kanzlei klären:

- Zeile 1090 (O14): `AVERAGE(F14:M14)` schließt die Vorjahressumme in Spalte F ein und rechnet gleitend; die übrigen Zeilen mitteln über G:N.
- Zeile 1100 (O19–R19): gleitender Mittelwert mit wechselnden Bereichen statt fester Basis.
- Zeile 1312, 1322, 1323: September (O) ohne Formel, Oktober bis Dezember mitteln über G:O.
- Zeile 1310: Juli (M33) ohne Zinsaufwand, alle anderen Monate rund 2.600.

**Austauschbare Einleseschicht**

Wie bei der Pillar-2-Pipeline kapselt ein eigenes Modul das Einlesen. Ändert sich das Quellformat, wird nur diese Schicht angepasst, nicht der Rest.

Umgesetzt in `prognosemodell/einlesen.py`: Blätter, die nicht im Kostenstellenformat sind (kein „Nr.“ in B4, keine Kostenstelle in B2, keine Basisjahrspalte), werden übersprungen und mit Grund gemeldet, etwa „Annahmen“. Werte stammen aus der Jahresspalte des Basisjahrs (im Muster S, inklusive Hochrechnung der offenen Monate). Kostenstellen ohne Stammdaten werden als neue Objekte angelegt, ihre steuerlichen Pflichtfelder bleiben leer. Abbruch mit Meldung nur bei Datenfehlern in einem Kostenstellenblatt (doppelte Kostenstelle, Formel ohne gespeicherten Wert, Text statt Zahl) oder wenn kein Blatt lesbar ist.

**Offene Punkte, nächste Woche in der Arbeit zu prüfen**

- [x] Gibt es ein Anlageverzeichnis mit Anschaffungskosten und Buchwerten je Objekt? Ja, DATEV-Export „Inventarübersicht“, Kostenstelle in KOST1 (Abschnitt 21).
- [x] Ist darin die Aufteilung Gebäude zu Grund und Boden schon enthalten? Ja, getrennte Anlagen je Konto. Grund und Boden trägt im Muster aber oft keine KOST1 und muss dann im Blatt Anlagen zugeordnet werden.
- [x] Haben die Kostenstellenblätter ein einheitliches Layout mit fester ObjektID? Muster liegt vor (B2 Kostenstelle, C2 Objekt), Einheitlichkeit über alle Blätter noch bestätigen.
- [x] Welcher Kontenrahmen (SKR03 oder SKR04) wird gebucht? Das Anlagenverzeichnis im Muster passt zu SKR04: 0235 Grundstückswerte, 0300–0360 Bauten, 0690 BGA, 0910 Finanzanlagen. Die Art je Konto ist in `einlesen.KONTEN_ART` hinterlegt und im Blatt änderbar.

## 17. Datenbedarf je Objekt

Vollständige Liste der Felder, die das Modell pro Objekt braucht, mit Quelle. Spalte Quelle sagt, ob das Feld automatisch aus den Kostenstellenblättern kommt, aus den Stammdaten gepflegt wird oder im Modell berechnet entsteht.

| Feld | Zweck | Quelle |
| --- | --- | --- |
| ObjektID | Schlüssel über alle Blätter | Stammdaten, feste Zelle je Kostenstellenblatt |
| Objektname | Lesbarkeit | Stammdaten oder Kostenstellenblatt |
| AK Gebäude | AfA-Basis, Buchwertfortschreibung | Stammdaten / Anlageverzeichnis |
| AK Grund und Boden | Gewinnaufteilung, keine AfA | Stammdaten / Anlageverzeichnis |
| Kaufjahr | Vorbesitzzeit § 6b, AfA-Start | Stammdaten / Anlageverzeichnis |
| AfA-Satz oder -Methode | Abschreibung je Jahr | Stammdaten, aus Baujahr abgeleitet |
| Restbuchwert Gebäude als Startwert | Startpunkt der Fortschreibung | Stammdaten / Anlageverzeichnis |
| Verkehrswert aktuell | stille Reserven, Verkaufsaufteilung | Stammdaten, Schätzung oder Gutachten |
| Miete im Basisjahr | Ertragsprognose | laufende Werte, Kostenstellenblatt |
| Erhaltungsaufwand im Basisjahr | Aufwandsprognose | laufende Werte, Kostenstellenblatt |
| weitere laufende Einnahmen | Ertragsprognose | laufende Werte, Kostenstellenblatt |
| weitere laufende Ausgaben | Aufwandsprognose | laufende Werte, Kostenstellenblatt |

**Nur bei geplantem Verkauf zusätzlich**

| Feld | Zweck | Quelle |
| --- | --- | --- |
| Verkaufsjahr | Zeitpunkt | Planung, Eingabe |
| Verkaufspreis oder Faktor | Erlös | Planung, Eingabe |
| Verkaufskosten | Gewinnminderung | Planung, Eingabe |
| 6b-Nutzung ja oder nein | Rücklage oder Sofortsteuer | Planung, Eingabe |

**Hinweis**

Die Trennung AK Gebäude zu Grund und Boden ist das kritischste Feld. Fehlt sie, lassen sich weder AfA noch § 6b sauber rechnen. Falls das Anlageverzeichnis sie nicht ausweist, muss der Kaufpreis nachträglich aufgeteilt werden, etwa nach Bodenrichtwert oder BMF-Arbeitshilfe.

Seit Abschnitt 20 ist nur noch Pflicht, was die Buchhaltung immer liefert: ObjektID und Miete. Alle übrigen Felder füllt eine Annahme, solange kein echter Wert vorliegt. Bei einem verkauften Objekt sind die steuerlichen Stammdaten und der Verkaufspreis kritisch: Sie bestimmen Gewinn und Rücklage, die Mappe markiert sie orange und warnt.

## 18. Zielstruktur: Ausgabe im DATEV-BWA-Format, Sonderbereich Verkauf und Kauf

Festgehalten nach Rückmeldung der Kanzlei (Oktober 2026), Grundlage ist die Datei „Planungsreferenz.xlsx“. Die Datei enthält Mandantenzahlen und liegt deshalb nicht im Repository. Die Struktur ist hier beschrieben und als Vorlage mit erfundenen Werten nachgebaut.

**Grundsatz**

- Die Datenquelle ist die DATEV-BWA-Kostenstellenblattsammlung (Form 01), ein Blatt je Kostenstelle. Das ist das Zielbild für das Kostenstellenformat.
- Die Ergebnisse sollen wieder in derselben Struktur stehen: je Kostenstelle ein Blatt mit denselben BWA-Zeilen, die Planjahre rechts neben den Ist-Werten, dazu ein Summenblatt „Alle Objekte“.
- Verkauf und Kauf werden als Sonderbereich geführt, jeweils mit einer Ergebnissicht und einer Detailsicht.

**Vorlage**

`python -m prognosemodell.vorlagen` schreibt `vorlagen/Kostenstellen_BWA_Vorlage.xlsx`: Summenblatt „Alle Objekte“ (B2 „KSt“) und die Blätter „KSt 1“ und „KSt 2“ mit erfundenen Werten. Miete und Erhaltung von KSt 1 entsprechen dem Testobjekt.

| Bereich | Inhalt |
| --- | --- |
| B2, C2 | Kostenstelle, Bezeichnung |
| Zeile 4 | B „Nr.“, C „Bezeichnung kurz“, F „Jahr 2024“, G „Jahr 2025“, H–S Monate 2026, T „Jahr 2026“, U „Plan 2027“, V–AN 2028–2046 |
| ab Zeile 6 | BWA-Zeilen 1010 bis 1380 in der Reihenfolge der Referenz, auch die Leerzeilen mit Nummer |
| U–AN | Planjahre, in der Vorlage leer; gelb = hier schreibt das Modell |

Das Prüfskript `pruefen_einlesen` liest die Vorlage ein: Basisjahr aus „Jahr 2026“, Summenblatt übersprungen, Planspalten als Jahr erkannt.

**Ausgabe in die Planspalten (umgesetzt, `prognosemodell/bwa.py`)**

Die Mappe enthält hinter den Kontrollblättern:
- das Summenblatt „Alle Objekte“ (B2 „KSt“);
- je Objekt und Neuobjekt ein BWA-Blatt im Layout der Vorlage;
- das Blatt „Verkauf und Kauf“.

**BWA-Blätter je Kostenstelle**

- Das Blatt heißt wie das eingelesene Kostenstellenblatt, sonst wie die ObjektID.
- B2 trägt die ObjektID, alle Formeln suchen über sie.
- Links stehen die Ist-Werte: Vorjahre, Monate und Basisjahr, wie eingelesen. Ohne eingelesene BWA zeigt das Basisjahr Miete, weitere Einnahmen, Erhaltung und weitere Ausgaben aus dem Objektblatt.
- Rechts stehen die Planjahre als Formeln aus der Prognose, Szenario A.
- Zeile 5 trägt das Planjahr als Zahl, nur als Hilfe für die Formeln.
- Ein Wert 0 bleibt leer, deshalb sind die Planspalten eines verkauften Objekts ab dem Folgejahr leer und die eines Neuobjekts bis zum Kaufjahr.
- Die Summenzeilen sind Formeln wie in der BWA: 1345 = 1300 − 1320 + 1330, 1353 = 1345 + 1351 − 1352, 1380 = 1353 − 1355.

| BWA-Nr. | Planwert |
| --- | --- |
| 1020 | Miete aus der Prognose |
| 1090 | weitere Einnahmen |
| 1100–1220 | weitere Ausgaben × Anteil der Kostenart im Basisjahr (Ist-Wert / weitere Ausgaben Basisjahr) |
| 1260 | weitere Ausgaben minus 1100–1220; ohne Aufteilung die ganzen weiteren Ausgaben |
| 1240 | AfA aus der Fortschreibung (Steuerbilanz) |
| 1250 | Erhaltung |
| 1323 | Veräußerungsgewinn im Verkaufsjahr, Auflösung und Zuschlag im Fristjahr |
| 1312 | Veräußerungsverlust und Einstellung in die § 6b-Rücklage im Verkaufsjahr |
| 1310, 1322 | nur im Summenblatt: Zins der Alternativanlage, negativ als Zinsaufwand; Darlehenszins ab Stufe 2 |
| 1355 | nur im Summenblatt: Steuer aus dem Blatt Liquidität, mit Verlustvortrag (die Steuer entsteht bei der GmbH, nicht je Kostenstelle) |
| 1051–1092, 1280–1380 | Summenformeln |

Die Tabelle zeigt die Standardzuordnung. Bestandswerte (1020–1260 der Bestandsobjekte) gehen fest auf ihre Zeile. Verkauf, Rücklage, Neuobjekte und Zins sind **Sonderposten**, ihre Zeile steuert das Blatt BWA-Zuordnung (unten). Standard wie in der Planungsreferenz: Verkauf und Rücklage im neutralen Ergebnis (1312, 1323), die außerordentlichen Zeilen 1351 und 1352 leer.

**Blatt BWA-Zuordnung (Steuerung und Kontrolle)**

- Je Sonderposten eine BWA-Nr. (gelb, Auswahlliste), daneben Bezeichnung der Zeile, Ertrag/Aufwand, Standard und die Summe über die Planjahre. Abweichung vom Standard ist dunkelgelb, eine ungültige Nr. rot.
- Zulässig sind die Einzelzeilen 1020, 1090, 1240, 1250, 1260, 1310, 1312, 1322, 1323, 1351, 1352. Die Kostenarten 1100–1220 bleiben den eingelesenen Kosten vorbehalten, Summenzeilen und Steuern sind ausgeschlossen.
- Jeder Posten ist ergebniswirksam gerechnet (Ertrag +, Aufwand −). Eine Aufwandszeile nimmt ihn mit umgekehrtem Vorzeichen auf. So bleibt das Ergebnis bei jeder Zuordnung gleich, nur der Ausweis ändert sich.
- „Verkauf ausweisen“: netto bucht nur Veräußerungsgewinn bzw. -verlust. Brutto bucht den Verkaufspreis als Ertrag sowie Verkaufskosten und Buchwertabgang (Gebäude + G+B) als Aufwand; die Summe ist derselbe Gewinn.
- Posten: Veräußerungsgewinn, -verlust (netto); Veräußerungspreis, -kosten, Buchwertabgang (brutto); Einstellung, Auflösung, Gewinnzuschlag § 6b; Neuobjekte Mieten, weitere Einnahmen, Erhaltung, weitere Ausgaben, Abschreibungen; Zinsertrag und Zinsaufwand (nur Summenblatt).
- Unter jeder BWA steht der Block „Herleitung Sonderposten“: je Posten die Ziel-Nr. (Spalte D) und der Betrag je Planjahr für diese Kostenstelle. Die Ziel-Nr. steht bewusst nicht in Spalte B, damit ein erneutes Einlesen die BWA-Zeilen nicht verwechselt.
- Kontrolle je Planjahr: Ergebnis lt. BWA Alle Objekte (1353) gegen Ergebnis vor Verlustvortrag im Blatt Liquidität, Differenz muss 0 sein. Die Prüfung „bwa\_zuordnung“ (Warnung) zählt ungültige Nummern und Jahre mit Differenz.
- Generator: Modell.bwa\_zuordnung, z. B. {"verkauf": "brutto", "erloes": 1351}.

Benannte Bereiche: zuo\_Verkauf, zuo\_<Posten> (BWA-Nr.), zuo\_Nr, zuo\_Nummern, zuo\_Liste, zuo\_Differenz; bwah\_<Posten> (Herleitung im Summenblatt über die Planjahre).

**Summenblatt „Alle Objekte“**

- Die Planwerte rechnen über die ganze Prognose, also auch über Objekte, die erst in Excel angelegt werden und kein eigenes Blatt haben.
- Die Kostenarten 1100–1220 sind die Summe der Kostenstellenblätter, 1260 nimmt den Rest der weiteren Ausgaben auf.
- Die Ist-Spalten sind die Summe der Kostenstellenblätter.
- Abgleich, im Prüfskript je Jahr geprüft:
  - Ergebnis vor Steuern (1345, bei Standardzuordnung) bzw. 1353 (bei jeder Zuordnung) = Ergebnis vor Verlustvortrag im Blatt Liquidität
  - Vorläufiges Ergebnis (1380) = Ergebnis nach Steuern im Blatt Auswertung
  - Abschreibungen (1240) = AfA im Blatt Liquidität

Benannte Bereiche: bwa\_<Nr.> je Zeile über die Planjahre, bwa\_Jahr.

**Sonderbereich Verkauf und Kauf**

Je Vorgang gibt es zwei Sichten. In der Referenz stehen sie rechts neben der BWA (ab Spalte AP); im Modell stehen sie im eigenen Blatt „Verkauf und Kauf“.

**Umsetzung im Blatt „Verkauf und Kauf“**

- Oben: Endvermögen nach latenter Steuer der Szenarien A, B, C und Baseline, mit Differenz zu A.
- Je Zeile des Verkaufsblatts ein Block, mindestens drei, damit in Excel ergänzte Verkäufe passen. Der Block liest den Verkauf per INDEX und die zugehörige Rücklage aus derselben Zeile des Rücklagenblatts.
- Neuobjekte gehören zum Verkauf, wenn ihre Quelle-RücklageID die Rücklage des Verkaufs nennt und ihr Status OK ist (Prognosespalte prg\_Quelle).
- Vergleichsjahr ist das erste volle Jahr nach Verkauf und letztem Kauf aus der Rücklage, höchstens das letzte Prognosejahr.
- **Ergebnissicht:**
  - Erlös, Reinvestition (Kaufpreis und Nebenkosten), Übertrag § 6b, Steuer auf den Gewinn ca. (Gewinn − Rücklage + Auflösung + Zuschlag) × Grenzsteuersatz.
  - Kapitalanlage = Nettoerlös − Reinvestition − Steuer ca.
  - Restschuld: Stufe 2.
  - Vergleich im Vergleichsjahr, Ausgangsfall (halten) gegen Alternative (Neuobjekte aus der Rücklage plus Zins auf die Kapitalanlage): Mietertrag, Kapitalertrag, Aufwand, vorläufiges Ergebnis, liquider Überschuss vor und nach Steuern ca., mit Differenz.
- **Detailsicht:**
  - Einzelauflistung des Verkaufs (gesamt, G+B, Gebäude): Anteil, Veräußerungspreis, Kosten, Buchwert, Gewinn, Rücklage, übertragen, Fristjahr, Auflösung, Zuschlag.
  - Planung im Vergleichsjahr: Mieten, weitere Einnahmen, Zinsertrag, Erhaltung, weitere Ausgaben, Abschreibungen, Zinsen und Tilgungen (Stufe 2, 0), vorläufiges Ergebnis, Cash Flow.
  - Ausgangsfall: die Werte des Objekts bei Halten, fortgeschrieben wie in der Prognose; die AfA aus dem Buchwert bei Halten.
- **Detailsicht Kauf:** je Neuobjekt eine Spalte: Kaufdaten, AK G+B und Gebäude, ü1 bis ü3, Übertrag gesamt, AfA-Bemessungsgrundlage, AfA-Methode und -Satz, im ersten vollen Jahr Mieten, Erhaltung, AfA, Ergebnis und Cash Flow.
- „Steuern ca.“ rechnen wie die Referenz: Ergebnis × Grenzsteuersatz, ohne Verlustvortrag. Die genaue Steuer je Jahr steht im Blatt Liquidität.
- Die Kostenaufstellung je Kostenart steht im BWA-Blatt der Kostenstelle. Der Sonderbereich zeigt Erhaltung und weitere Ausgaben zusammengefasst, wie die Planung der Referenz.

*Ergebnissicht („Für Berichterstattung“), je Vorgang eine Spalte:*

| Zeile | Inhalt |
| --- | --- |
| Bewertung/Erlös | Verkaufspreis |
| Reinvestition | Kaufpreis des Neuobjekts |
| Kapitalanlage | freie Mittel in der Alternativanlage |
| Übertrag § 6b EStG | übertragene Rücklage (negativ) |
| Restschuld aktuell / nach Umstrukturierung | Darlehen vor dem Verkauf, Finanzierung danach |
| Vergleich Ausgangsfall gegen Alternative | Mietertrag netto, Kapitalertrag, Aufwand (Nebenkosten neutralisiert), vorläufiges Ergebnis, liquider Überschuss vor Steuern, Steuern ca., liquider Überschuss nach Steuern |

*Detailsicht („Einzelauflistung“), je verkaufter bzw. gekaufter Kostenstelle:*

- Veräußerungspreis, aufgeteilt nach G+B und Gebäude (Anteil in Prozent), abzüglich Veräußerungskosten und Buchwert je Teil, ergibt den Veräußerungsgewinn je Teil.
- Neue Mittel: Mietertrag des Neuobjekts (Rendite auf den Kaufpreis), Eigenkapital (= Buchwert-Rückfluss), abzüglich Restschuld, ergibt die Anlage mit ihrer Rendite.
- Reinvestition: Kaufpreis, Übertrag § 6b, AfA-Bemessungsgrundlage, AfA Steuerbilanz, AfA Handelsbilanz, Finanzierung mit Zins und Tilgung.
- Planung Ausgangsfall gegen Fall, je Handels- und Steuerbilanz: Mieten, Mietnebenkosten, Zinsertrag, Aufwendungen, Abschreibungen, Zinsen, vorläufiges Ergebnis, Tilgungen, AfA zurück, Cash Flow. Die Kostenzeilen des Ausgangsfalls kommen aus dem BWA-Blatt der verkauften Kostenstelle.

**Rechenwege der Referenz (nachgerechnet)**

| Größe | Rechnung in der Referenz |
| --- | --- |
| Aufteilung Erlös | G+B 40 %, Gebäude 60 % des Veräußerungspreises |
| Veräußerungsgewinn | je Teil Erlösanteil − Buchwert; gesamt Preis − Buchwert |
| Eigenkapital | = Buchwert-Rückfluss (Preis − Gewinn) |
| Anlage | Eigenkapital − Restschuld; Rendite 3 % |
| Mietertrag Neuobjekt | Kaufpreis × 3 % |
| Aufwendungen Neuobjekt | 5 % des Mietertrags |
| AfA-Bemessungsgrundlage | Reinvestition − Übertrag § 6b (Übertrag auf die gesamten AK, ohne Trennung Gebäude/G+B) |
| Finanzierung | = AfA-Bemessungsgrundlage, also der Teil der Reinvestition, den der Gewinn nicht deckt; Zins 3,5 %, Tilgung 2 % |
| AfA Steuerbilanz | AfA-Bemessungsgrundlage × 5 % degressiv |
| AfA Handelsbilanz | 75 % der AK (geschätzter Gebäudeanteil) × 2 % linear, ohne § 6b-Kürzung |
| Aufwand neutralisiert | Aufwendungen + Mietnebenkosten (Umlagen gegen Aufwand) + AfA + Zinsen |
| Cash Flow | vorläufiges Ergebnis − Tilgung + AfA |
| Steuern ca. Ausgangsfall | 45 % des vorläufigen Ergebnisses |

**Abweichungen zum bisherigen Modell, offen zu klären**

- [x] **Handels- und Steuerbilanz:** Entschieden: im ersten Durchlauf nur die Steuerbilanz. Steuer, Liquidität und Endvermögen hängen nur an ihr. Die Handelsbilanz ändert das ausgewiesene Jahresergebnis (keine § 6b-Kürzung, höhere AfA) und damit das Ausschüttungspotenzial und die latenten Steuern im Jahresabschluss. Sie kann später als zweite AfA-Spalte dazukommen.
- [x] **Degressive AfA 5 %** für das Neuobjekt: umgesetzt als AfA-Methode im Blatt Neuobjekte (Abschnitt 13), mit Wechsel zur linearen AfA.
- [x] **Übertrag auf die gesamten AK:** Entschieden: das Modell bleibt bei der Trennung. Die Referenz kürzt die AfA-Basis um den ganzen Übertrag. Das Modell trennt Gebäude- und G+B-Gewinn nach § 6b Abs. 1 (G+B-Gewinn zuerst auf G+B). Ergebnis gleich, solange G+B-Gewinn auf das Gebäude passt; die AfA-Basis unterscheidet sich, wenn das Neuobjekt einen G+B-Anteil hat.
- [x] **Finanzierung und Restschuld:** Entschieden: bleibt in Stufe 2, dort mit Vorrang. In der Referenz ist es schon Teil der Rechnung (Ablösung der Restschuld, Darlehen für den nicht gedeckten Teil). Bis dahin zeigt der Sonderbereich Restschuld und Finanzierung nicht.
- [ ] **Mietnebenkosten:** Die Referenz neutralisiert Umlagen gegen den Aufwand. Im Modell zählt BWA 1020 komplett als Miete. Klären, auf welchem Konto oder welcher BWA-Zeile die Umlagen stehen.
- [ ] **Steuersatz:** Die Referenz rechnet mit etwa 45 %, das Parameterblatt mit 30 % (GmbH ohne erweiterte Kürzung wäre rund 30 %). Die 45 % der Alternative lassen sich nicht aus den angezeigten Werten herleiten.
- [x] **Zeitpunkt:** Die Referenz vergleicht ein Jahr (Basis 2025) statisch; das Modell rechnet 20 Jahre. Umgesetzt: Die Ergebnissicht zeigt das erste volle Jahr nach Verkauf und Kauf, oben im Blatt das Endvermögen nach 20 Jahren je Szenario.
- [ ] **Rückführung:** Die BWA-Blätter liegen in der Prognosemappe und rechnen in Formeln. Ein Export als eigene Datei mit festen Werten, etwa zum Einspielen in die Kanzlei-Excel, ist noch offen.

## 19. VBA-Steuerung und Plausibilitätsprüfungen (Etappe 9)

**Grundsatz:** Geprüft wird in Formeln, nicht in VBA. Das Blatt Prüfung rechnet jede Prüfung als Formel. So bleiben Fehler auch ohne Makros sichtbar, etwa in der .xlsx-Fassung oder bei deaktivierten Makros. Die Makros lesen das Ergebnis nur, schreiben Eingabewerte und lösen die Neuberechnung aus.

**Blatt Prüfung**

Je Prüfung eine Zeile mit Art, Anzahl betroffener Zeilen (pr\_Anzahl) und Ergebnis (pr\_Ergebnis, „OK“ oder die Art). Darüber stehen das Gesamtergebnis (pr\_Gesamt) sowie die Anzahl der Fehler und Warnungen (pr\_Fehler, pr\_Warnungen). Das Gesamtergebnis steht auch auf dem Parameterblatt (par\_StatusPruefung) und in der Übersicht (ueb\_Pruefung), rot, solange es nicht „OK“ lautet.

| Prüfung | Art | Formel (fachlich) |
| --- | --- | --- |
| Objekte mit Status ungleich OK | Fehler | Statusspalte Objekte |
| Verkäufe mit Status ungleich OK | Fehler | Statusspalte Verkäufe, ohne „§ 6b unzulässig“ |
| Neuobjekte mit Status ungleich OK, darunter Fristverstoß | Fehler | Statusspalte Neuobjekte |
| Steuerwelt außerhalb des MVP | Fehler | par\_Steuerwelt ≠ GmbH |
| Vorbesitzzeit für § 6b zu kurz | Warnung | Status „§ 6b unzulässig: Vorbesitzzeit zu kurz“ |
| Rücklage nur teilweise übertragen | Warnung | Auflösung > 0 und ein Neuobjekt nennt die Rücklage |
| Rücklage ohne Neuobjekt | Warnung | Auflösung > 0 und kein Neuobjekt nennt sie |
| Drei-Objekt-Grenze | Warnung | mehr als par\_DOGrenze gültige Verkäufe im Zeitraum par\_DOJahre, der mit einem Verkaufsjahr endet |
| Frist endet nach Prognoseende | Hinweis | Hinweis im Rücklagenblatt gesetzt |
| Liquidität negativ | Hinweis | Jahre mit liq\_Kum < 0 (Finanzierungsbedarf, Stufe 2) |
| Objekte ohne Verkehrswert | Hinweis | Status OK und Verkehrswert leer |

- Fehler: Eingabe unvollständig oder unzulässig, die Zeile rechnet nicht oder nur teilweise mit.
- Warnung: rechnet, ist aber steuerlich ungünstig oder fachlich zu prüfen.
- Hinweis: zur Kenntnis, zählt nicht ins Gesamtergebnis.

„Gebäudeanteil des Neuobjekts ≥ übertragener Gebäudegewinn“ aus Abschnitt 5 steckt in der Warnung zum Teilübertrag: Reicht der Gebäudeanteil nicht, bleibt ein Rest, der im Fristjahr aufgelöst wird. „G+B plus Gebäude = Kaufpreis“ ist durch den Aufbau erfüllt: Beim Bestandsobjekt werden beide AK getrennt erfasst, beim Neuobjekt ergibt der Anteil G+B mit dem Rest Gebäude immer den Kaufpreis.

Zur Drei-Objekt-Grenze: Die Prüfung ist vereinfacht. Sie zählt Verkäufe im Zeitraum und nicht Verkäufe innerhalb von fünf Jahren nach dem Erwerb. Bei der GmbH sind die Einkünfte ohnehin gewerblich. Die Gefahr ist, dass die Objekte als Umlaufvermögen gelten und § 6b dann entfällt. Fachlich prüfen.

**VBA-Module**

Die Module liegen als Quelltext in `prognosemodell/vba`, ohne Umlaute; Meldungstexte laufen über `Txt()`.

| Modul | Makros |
| --- | --- |
| modStart | Schaltflächen auf dem Parameterblatt anlegen (beim Öffnen), Hilfen `Bereich`, `Txt`, `Euro` |
| modRechnen | `NeuBerechnen` (vollständig), `NeuBerechnenStarten` |
| modPruefung | `AnzahlFehler`, `AnzahlWarnungen`, `Auffaelligkeiten`, `PruefungStarten`, `PruefungBestanden` (fragt bei Fehlern nach) |
| modObjekte | `ObjektAnlegen`, `ObjektDuplizieren`, `ObjektEntfernen`, `LeereBloeckeAusblenden` und die Bedienung per Schaltfläche |
| modVarianten | `VarianteFesthalten`, `VariantenLeeren`, `FreieVariante` |
| ThisWorkbook | beim Öffnen `Einrichten` |

Die Makros fügen nie Zeilen ein und löschen nie welche. Sie schreiben nur in die gelben Eingabefelder (obj\_Eingabe). Jede Objektzeile hat in der Prognose ihren festen Block, deshalb rechnet ein angelegtes oder dupliziertes Objekt sofort ohne Formelbruch mit. „Objekt entfernen“ leert die Eingaben. „Leere Prognoseblöcke ausblenden“ blendet die Blöcke ohne ObjektID aus, auch die leeren Neuobjektblöcke.

**Einbetten:** openpyxl kann kein VBA-Projekt erzeugen, nur ein vorhandenes übernehmen. `prognosemodell/makros.py` baut deshalb per LibreOffice (headless, UNO) aus den Quelltexten ein VBA-Projekt in einer leeren Hilfsmappe mit denselben Codenamen. openpyxl übernimmt daraus nur `vbaProject.bin`. Aufruf: `python -m prognosemodell --makros`. LibreOffice wird nur beim Bauen gebraucht, die fertige .xlsm nur Excel. Ohne `--makros` entsteht wie bisher eine .xlsx ohne Makros, mit denselben Formeln und Prüfungen.

**Abnahme Etappe 9** (im Prüfskript)

- Jede Prüfung schlägt bei einem eingebauten Fehler an, mit der richtigen Anzahl und Art. Ein sauberes Testobjekt ergibt „OK“ in Prüfung, Parameterblatt und Übersicht.
- Objekt anlegen ohne Formelbruch. Eine Kopie von OBJ-001 rechnet im eigenen Block (Miete 2027 61.200, Buchwert 2046 160.000), die Übersicht zählt sie mit.
- Eine doppelte ID oder eine fehlende Quelle wird abgewiesen, ohne dass etwas geschrieben wird. Eine entfernte Zeile wird wieder belegt.
- Leere Blöcke lassen sich aus- und einblenden.
- Eine festgehaltene Variante enthält dieselben Werte wie das Blatt Vergleich.

Die Makrofälle laufen im Prüfskript in LibreOffice. Dort bricht `Err.Raise` die Funktion ab und liefert 0, Excel zeigt stattdessen die Meldung. Die Schaltflächen und Eingabedialoge sind nur in Excel bedienbar und werden nicht automatisch geprüft.

## 20. Datenlage: Auffülllogik, Farben, Startblatt und Schnellcheck

Anlass: Zu vielen Bestandsobjekten liegen nur die laufenden Buchungen vor, keine Verkehrswerte und keine Kaufdaten. Das Modell soll trotzdem rechnen und zeigen, wie belastbar das Ergebnis ist. Dazu fährt es zweigleisig: Wo fehlende Daten nur Feinschliff sind, füllt eine plausible Annahme; wo sie das Ergebnis bestimmen, markiert die Mappe sie und warnt.

**Auffülllogik**

Die Annahme steht als Formel direkt in der leeren Eingabezelle (blau). Wer einen echten Wert eintippt, ersetzt sie (gelb). Die Formeln bauen aufeinander auf, die Sätze stehen zentral auf dem Parameterblatt im Abschnitt „Annahmen bei fehlenden Daten“.

| Feld | Annahme | Parameter |
| --- | --- | --- |
| Erhaltung | Miete × Quote | par\_AnnErhQuote (10 %) |
| Baujahr, Großmaßnahme | siehe „Erhaltung nach Gebäudealter“ unten | |
| Verkehrswert | (Miete + weitere Einnahmen) × Vervielfältiger | par\_AnnVervielfaeltiger (20) |
| Verkehrswertanteil Gebäude | Gebäudeanteil | par\_AnnGebaeudeanteil (75 %) |
| AfA-Satz | AfA-Satz Bestand | par\_AnnAfASatz (2 %) |
| Kaufjahr | Basisjahr − Jahre seit Kauf | par\_AnnHaltedauer (15) |
| AK Gebäude | AfA lt. Buchhaltung / AfA-Satz; ohne AfA: Verkehrswert × Gebäudeanteil / (1 + Wertsteigerung)^(Basisjahr − Kaufjahr) | |
| AK G+B | AK Gebäude × (1 − Gebäudeanteil) / Gebäudeanteil | |
| Restbuchwert | AK Gebäude × (1 − (Basisjahr − Kaufjahr + 1) × AfA-Satz), mindestens 0; 0, wenn die AfA lt. Buchhaltung 0 ist (abgeschrieben) | |
| AfA je Jahr (Prognose) | AfA lt. Buchhaltung, auch 0; ohne sie AK Gebäude × AfA-Satz | |
| Verkaufspreis | Verkehrswert × (1 + Wertsteigerung)^(Verkaufsjahr − Basisjahr) | |

Neues Eingabefeld „AfA Basisjahr lt. Buchhaltung“ (BWA 1240, obj\_AfABWA): Mit ihm trifft die AfA der Prognose die Buchhaltung, auch wenn AK und Kaufjahr fehlen.

**AfA der Buchhaltung steuert die Prognose** (Feld „AfA je Jahr (Prognose)“, obj\_AfAJahr): Die Prognose schreibt die AfA des Basisjahrs fort, bis der Restbuchwert verbraucht ist; danach 0. Zeigt die Buchhaltung keine AfA (BWA 1240 leer oder 0 im eingelesenen Blatt), gilt das Gebäude als abgeschrieben: AfA 0 und Restbuchwert 0, der ganze Gebäudeerlös ist dann Gewinn. Ohne eingelesene AfA bleibt es bei AK Gebäude × AfA-Satz. Ein eingetippter Wert ersetzt die Annahme, etwa wenn die AfA des Basisjahrs eine Sonder- oder Teil-AfA enthält. Das Feld ist blau, zählt aber nicht als eigene Annahme: es leitet nur aus AfA lt. Buchhaltung bzw. AK und Satz ab, deren Annahmen schon zählen. Vorher wirkte die AfA lt. Buchhaltung nur über die geschätzten AK Gebäude und ging verloren, sobald AK Gebäude eingetragen oder die AfA 0 war. Weitere Einnahmen und Ausgaben bleiben leer = 0. Pflicht sind nur ObjektID und Miete.

Wird eine Annahme gelöscht, ohne einen Wert einzutragen, meldet der Status „Wert fehlt: Annahme gelöscht“. Das Makro „Annahmen wiederherstellen“ füllt leere Felder aus der ausgeblendeten Vorlagezeile (obj\_Vorlage). „Objekt entfernen“ stellt die Annahmen der Zeile selbst wieder her, „Objekt duplizieren“ kopiert Formeln als Formeln.

**Farblogik der Eingabezellen** (bedingte Formatierung, erste passende Regel gilt)

| Farbe | Bedeutung | Regel |
| --- | --- | --- |
| rot | Pflichtwert fehlt oder Annahme gelöscht | ObjektID gesetzt, Zelle leer |
| orange | kritische Annahme | Formel in der Zelle, Objekt hat einen Verkauf; Felder Verkehrswertanteil, AfA-Satz, Kaufjahr, AK, Restbuchwert sowie der Verkaufspreis |
| blau | Annahme | Formel in der Zelle (ISTFORMEL) |
| grün | aus der Buchhaltung eingelesen | Wert gleich dem eingelesenen Wert in den ausgeblendeten Spalten rechts |
| gelb | händisch eingetragen | alle übrigen Eingaben |

Je Objekt zählen die Spalten „Annahmen (blau)“ und „kritische Annahmen (orange)“ mit (obj\_Annahmen, obj\_Kritisch), je Verkauf „Preis angenommen“ (vk\_PreisAnnahme). Das Blatt Prüfung meldet:
- **Warnung:** Verkauf mit kritischen Annahmen.
- **Hinweis:** Objekte mit Annahmen. Er ersetzt den früheren Hinweis „ohne Verkehrswert“.

**Erhaltung nach Gebäudealter und Großmaßnahmen**

Anlass: Halten lag im Abnahmefall vorn, unter anderem weil das Neuobjekt mehr Erhaltung trug (1 % des Kaufpreises, 20 % der Miete) als das 20 Jahre alte Objekt (13 % der Miete). Alte Gebäude werden mit den Jahren teurer, Neubauten sind anfangs fast wartungsfrei, und große Einzelmaßnahmen fallen nur beim Halten an.

| Baustein | Rechnung | Parameter (Standard) |
| --- | --- | --- |
| Alterung Bestand | Erhaltung × (1 + Alterung)^(Jahre über dem Schwellenalter seit dem Basisjahr), zusätzlich zur Erhaltungssteigerung | par\_ErhAlterungAb (30), par\_ErhAlterung (1,5 %) |
| Baujahr | Eingabe; leer: Basisjahr − Gebäudealter | par\_AnnGebaeudealter (40) |
| Großmaßnahme Jahr | Eingabe (0 = keine); leer: Baujahr + Alter, bei schon älteren Gebäuden erstes Prognosejahr + Vorlauf; jenseits des Rasters 0 | par\_SanAlter (50), par\_SanVorlauf (2) |
| Großmaßnahme Betrag | Eingabe in heutigen Preisen; leer: Verkehrswert × Gebäudeanteil × Quote; wächst mit der Erhaltungssteigerung | par\_SanQuote (15 %, 0 % schaltet die Annahme ab) |
| Neuobjekt Anlauf | Erhaltung in den ersten Jahren nach dem Kauf × Faktor | par\_NeuErhAnlaufJahre (10), par\_NeuErhAnlaufFaktor (50 %) |
| Erhaltung Neuobjekt | Standard der Annahme gesenkt | par\_AnnNeuErhQuote 0,5 % statt 1 % |

- **Großmaßnahme als Erhaltungsaufwand:** Sie zählt als sofort abziehbarer Erhaltungsaufwand im Jahr der Maßnahme (BWA 1250).
  - Ob sie Herstellungskosten sind (Standardhebung) oder anschaffungsnaher Aufwand, ist fachlich zu prüfen. Aktivierung mit AfA ist nicht abgebildet.
  - Höhere Miete oder ein höherer Wert nach der Maßnahme sind nicht abgebildet.
- **Neue Prognosespalte prg\_ErhaltungHalten:** die Erhaltung, als würde nie verkauft. Die Baseline und der Ausgangsfall im Sonderbereich rechnen mit ihr, der Plan nur bis zum Verkaufsjahr. Ein Verkauf vor der Großmaßnahme erspart sie also dem Plan, nicht dem Halten.
- **Prüfskript:** Fälle vor dieser Logik rechnen ohne Alterung, Anlaufminderung und Großmaßnahmen (OHNE\_ALTERUNG), damit ihre Handrechnungen gelten. Das Testobjekt hat Baujahr 2007 und keine Großmaßnahme.

Wirkung im Abnahmefall (Endvermögen 2046, § 6b-Kette minus Halten):

| Variante | vorher | mit Alterslogik |
| --- | --- | --- |
| wie erfasst (Erhaltung neu 1 %) | −154.246 | −85.479 |
| Erhaltung neu 0,5 % | −40.167 | −2.592 |
| Altobjekt Baujahr 1975, Großmaßnahme 2029, Erhaltung neu 0,5 % | | +135.586 |

**Reinvestition aus dem Verkauf**

Neues Feld im Blatt Verkäufe: „reinvestieren“ (ja/nein). Bei ja entsteht in derselben Zeile des Blatts Neuobjekte ein Neuobjekt aus Formeln (blau):

| Feld | Wert |
| --- | --- |
| NeuID | „NEU-“ & ObjektID |
| Kaufjahr | Verkaufsjahr + par\_AnnReinvestJahre |
| Kaufpreis | Nettoerlös × par\_AnnReinvestQuote / (1 + par\_AnnNeuNebenkosten) |
| Nebenkosten | Kaufpreis × par\_AnnNeuNebenkosten |
| Anteil G+B, AfA-Satz, AfA-Methode, Mietrendite, Erhaltungsquote | Annahmen des Parameterblatts |
| Quelle | Rücklage des Verkaufs, wenn § 6b = ja |

Ein Neuobjekt aus dem Modell oder von Hand ersetzt die Zeile.

**Startblatt**

Das Startblatt ist das erste Blatt beider Mappen.
- **Handlungsempfehlung:** die beste Option nach Endvermögen nach latenter Steuer, mit dem Vorsprung gegenüber Halten. Ohne Verkauf erscheint ein Hinweis, was einzutragen ist.
- **Endvermögen je Option:** Halten, § 6b-Kette (Plan), sofort versteuern und reinvestieren, sofort versteuern und anlegen; mit Differenz zu Halten und Rang.
- **Kennzahlen:** Wert der § 6b-Kette (A − C), Steuer gesamt, tiefster Liquiditätsstand mit Jahr.
- **Belastbarkeit:**
  - „Nicht belastbar“ bei Fehlern.
  - „Vorläufig“ bei kritischen Annahmen.
  - Sonst „Belastbar im Rahmen der zentralen Annahmen“.
- Dazu Datenlage, Anleitung mit Links, Farblegende und die zentralen Annahmen.

Die Blattreiter sind nach Typ gefärbt: gelb Eingabe, grau Rechnung, blau Ausgabe, grün Kontrolle.

Benannte Bereiche: start\_Empfehlung, start\_Vorsprung, start\_Belastbarkeit, start\_Kritisch, start\_Optionen, start\_Werte.

**Schnellcheck**

`python -m prognosemodell --schnellcheck` (mit `--kostenstellen` aus der BWA) erzeugt `ausgabe/Schnellcheck_VV.xlsx`. Die Rechenlogik ist dieselbe, nur Eingabe und Ansicht sind schlanker:
- Im Blatt Verkäufe stehen „§ 6b nutzen“ und „reinvestieren“ als Annahme auf ja.
- Im Blatt Objekte sind die steuerlichen Stammdaten zu einer zugeklappten Spaltengruppe zusammengefasst (+ am Spaltenkopf).
- Prognose, Rücklagen, Liquidität, Auswertung und die BWA-Blätter sind ausgeblendet. Rechtsklick auf einen Reiter, „Einblenden“, holt sie zurück.

Mindesteingabe sind ObjektID und Miete je Objekt, besser auch Erhaltung und AfA aus der BWA, dazu je geplantem Verkauf Objekt und Jahr.

**Diagramme**

Alle Liniendiagramme haben dieselbe Achsenformatierung, damit sich in Excel nichts überlagert:
- Beträge in Tsd. €, die Einheit im Titel statt eines Achsentitels.
- Jahre schräg gestellt und immer am unteren Rand, auch bei negativen Werten.
- Titel und Legende außerhalb der Zeichenfläche.

**Prüfung im Prüfskript:**
- ein Objekt nur mit Miete, mit und ohne AfA lt. Buchhaltung, gegen die Handrechnung;
- AfA lt. Buchhaltung abweichend von AK × Satz und 0 (abgeschrieben), Auslaufen am Restbuchwert;
- BWA-Zuordnung brutto mit außerordentlichen Zeilen und mit ungültiger Nr.;
- ein Verkauf ohne Preis mit automatischer Reinvestition;
- die Annahmen des Schnellchecks;
- das Wiederherstellen per Makro;
- die Empfehlung im Abnahmefall.

**Offen:**
- Die Annahmesätze sind Platzhalter und mit der Kanzlei abzustimmen.
- Die Schaltflächen und die Farben sind in Excel zu sichten; das Prüfskript sieht nur die Werte.

## 21. Anlagenverzeichnis und Aufschlüsselung der Abschreibungen

Anlass: Zwei neue Datenquellen machen AK, Buchwert, Kaufjahr und AfA ohne Annahmen fest.
- **Anlagenverzeichnis:** DATEV-Export „Inventarübersicht“, eine Zeile je Anlage. Die Spalte KOST1 verweist auf die Kostenstelle, also auf das Objekt.
- **Kostenstellenblätter:** Die Kanzlei hat sie erweitert. Unter der BWA schlüsseln sie die Abschreibungen je Anlagengruppe auf, mit dem Buchwert zum Jahresende und der Jahres-AfA.

Vorher lief die AfA eines Objekts als ein Betrag (BWA 1240) weiter, bis der Restbuchwert verbraucht war. Laufen Anlagen verschieden lang (Außenanlagen 10 %, Gebäude 2 %), war das zu grob, und AK, Kaufjahr und Restbuchwert blieben Annahmen.

**Einlesen** (`einlesen.py`)

`--inventar PFAD` liest das Anlagenverzeichnis. Die Spalten werden über die Kopfzeile gesucht, Pflicht sind Inventar, Buchw. Wj-Ende und KOST1. Daraus wird:

| Feld im Blatt Anlagen | Quelle |
| --- | --- |
| Inventar-Nr., Bezeichnung, Konto, KOST1, AHK-Datum, AHK | gleichnamige Spalten (AHK Wj-Ende) |
| Buchwert Stand | Buchw. Wj-Ende |
| AfA-Satz | AfA-% / 100 |
| AfA im Stand-Jahr | N-AfA und S-Abschr., jeweils Wj-Ende − Wj-Beginn (nur Kontrolle) |
| Art | AfA-Art Lin.Geb. = Gebäude, Anlag./Bau = im Bau, Finanzanl. = Finanzanlage; sonst nach Konto (SKR04): 200–239 G+B, 240–399 Gebäude, 400–699 BGA, 700–799 im Bau, 800–999 Finanzanlage |
| Methode | Keine AfA, Anlag./Bau, Finanzanl. oder ohne Satz: keine; Geom.degr.: degressiv; sonst linear |
| ObjektID | ObjektID im Blatt Objekte mit derselben Endnummer wie KOST1, sonst „KSt <KOST1>“; ohne KOST1 leer |

Abgegangene Anlagen (Datum in Abgang) entfallen. Abbruch mit Meldung bei doppelter Inventar-Nr., fehlender Kopfzeile oder Text statt Zahl.

Das Jahr der Buchwerte (Wj-Ende) steht auf dem Parameterblatt (par\_AnlStand), Standard Basisjahr − 1. `--inventar-stand` setzt es; sonst gilt das Jahr im Dateinamen (Inventar\_2025.xlsx).

`--kostenstellen` liest zusätzlich die Aufschlüsselung unter der BWA. Sie steht in Spalte C ohne BWA-Nr. und hat drei Blöcke:
- „Buchwert, JE“: eine Zeile je Gruppe;
- „Abschreibungen JW“: die Jahres-AfA je Gruppe;
- „Abschreibungen MW“: beendet die Blöcke.

Gelesen wird die Spalte des Jahres par\_AnlStand, im Muster F (2025). Buchwert und AfA einer Gruppe werden über die Beschriftung verbunden: bei gleicher Beschriftung, sonst wenn die eine mit der anderen beginnt („TG“ zu „TG 24“), sonst nach der Reihenfolge. Zeilen ohne Beschriftung (Zwischensummen) und Gruppen ohne Buchwert und AfA entfallen.

Jede Gruppe wird eine Anlage der Art Gebäude mit Buchwert und AfA p. a., aber ohne AHK und Datum. Sie zählt nur für Kostenstellen ohne Anlage im Anlagenverzeichnis. Liegen beide Quellen vor, gibt das Einlesen je Kostenstelle beide Buchwerte aus; im Muster stimmen sie überein (KSt 1: 971.005 €).

**Blatt Anlagen**

Eine Zeile je Anlage. Die Eingaben sind gelb, AfA p. a. steht als Formel darin (blau), und die berechneten Spalten sind grau:

| Spalte | Rechnung |
| --- | --- |
| AfA p. a. | linear: AHK × Satz, auf volle Euro aufgerundet (ROUNDUP(ROUND(…;2);0)), trifft die DATEV-AfA im Muster; degressiv: Buchwert Stand × Satz; keine: 0. Eintippen ersetzt die Formel |
| Gruppe | G+B; Gebäude, BGA, im Bau = abnutzbar; sonst leer (nicht im Modell) |
| Zugangsjahr | Jahr des AHK-Datums, nur G+B und Gebäude |
| Buchwert Ende Basisjahr | linear MAX(Buchwert Stand − AfA p. a. × (Basisjahr − Stand); 0), degressiv Buchwert × (1 − Satz)^(Basisjahr − Stand) |
| AfA Basisjahr, AfA je Prognosejahr t | linear MIN(AfA p. a.; MAX(Buchwert Stand − AfA p. a. × (t − 1 − Stand); 0)), degressiv Buchwert × (1 − Satz)^(t − 1 − Stand) × Satz; 0 bis zum Stand. Ist der Stand das Basisjahr, gilt als AfA Basisjahr die AfA lt. Inventar |
| Status | Pflichtfeld fehlt (Inventar-Nr., Art, Buchwert, Methode), Inventar-Nr. doppelt, nicht im Modell (Art), ohne ObjektID, ObjektID fehlt im Blatt Objekte, sonst OK |

Jede Zelle rechnet in geschlossener Form für sich, ohne Kette über die Jahre. Benannte Bereiche: anl\_Nr, anl\_ID, anl\_Art, anl\_AHK, anl\_BWStand, anl\_AfA, anl\_Gruppe, anl\_Zugang, anl\_BWBasis, anl\_AfABasis, anl\_Status, dazu anl\_AfAJahre über alle Jahresspalten.

**Objekte**

Neue Hilfsspalten nach den Annahmen zählen je Objekt die Anlagen mit Status OK. Daneben stehen die AfA Basisjahr lt. Anlagen und ihre Abweichung zu BWA 1240 (obj\_AnlAbn, obj\_AnlAK, obj\_AnlGuB, obj\_AnlKauf, obj\_AnlAfA, obj\_AnlDiff). Die Annahmeformeln der Eingabezellen fragen zuerst das Blatt Anlagen:

| Feld | aus dem Blatt Anlagen, wenn | Wert |
| --- | --- | --- |
| AK Gebäude, AfA-Satz | abnutzbare Anlagen da, alle mit AHK | Summe AHK; Satz = Summe AfA p. a. / AK |
| AK G+B | Anlage G+B da | Summe Buchwert G+B Ende Basisjahr |
| Kaufjahr | Zugangsjahr da | frühestes Zugangsjahr (MINIFS); spätere Zugänge gelten als nachträgliche AK |
| Restbuchwert | abnutzbare Anlagen da | Summe Buchwert Ende Basisjahr |
| AfA je Jahr | abnutzbare Anlagen da | AfA im ersten Prognosejahr (Anzeige) |

Kommt ein Wert aus dem Blatt Anlagen, ist die Zelle grün und zählt nicht als Annahme (auch nicht als kritische). Ohne Anlagen bleibt die bisherige Annahme. Gruppen aus dem Kostenstellenblatt haben keine AHK. Dann kommen nur Restbuchwert und AfA aus dem Blatt, AK und Kaufjahr bleiben Annahmen.

Restbuchwert und AfA umfassen auch BGA und Anlagen im Bau der Kostenstelle, so wie BWA 1240 alle Abschreibungen der Kostenstelle zeigt. Beim Verkauf gehen sie mit dem Gebäudebuchwert ab.

**Prognose**

Neue Spalte prg\_AfAHalten („AfA bei Halten“). Mit Anlagen ist sie die Summe der AfA je Anlage im Jahr, sonst AfA je Jahr; höchstens der Buchwert bei Halten des Vorjahrs. Darauf bauen auf:
- Buchwert bei Halten = Vorjahr − AfA bei Halten;
- AfA im Plan = AfA bei Halten × aktiv.

Ohne Anlagen ergibt das dieselben Werte wie vorher.

**Ausgabe und Prüfung**

- Jedes BWA-Blatt einer Kostenstelle mit Anlagen zeigt unter der Herleitung die Blöcke „Buchwert, JE“ und „Abschreibungen JW“ je Anlage. Spalten: Vorjahr = Stand, Basisjahr, Planjahre. Die Inventar-Nr. steht in Spalte D. Die Werte gelten bei Halten.
- Prüfung, Hinweis: Anlagen G+B, Gebäude, BGA oder im Bau ohne Objekt oder unvollständig.
- Prüfung, Hinweis: AfA lt. Anlagen weicht um mehr als 1 € von BWA 1240 ab.
- Startblatt: Objekte mit Anlagen und Anlagen ohne Objekt.

**Muster der Kanzlei (Inventar 2025, KSt 1)**

- 204 Anlagen, eine abgegangen. KSt 1 hat sechs abnutzbare Anlagen: TG 305001, Außenanlagen 306001, 310001 und 311001 (abgeschrieben), Wohnbauten 360010 und 360012.
- Restbuchwert Ende 2026: 928.630 €. AfA 42.375 € je Jahr, gleich BWA 1240. Buchwert Ende 2046: 81.130 €. Das trifft die Fortschreibung im Kostenstellenblatt (S50, AM50).
- Grund und Boden 200001 (30.12.1998) trägt keine KOST1. AK G+B von KSt 1 bleibt daher eine Annahme, bis die ObjektID im Blatt Anlagen eingetragen ist.

**Prüfung im Prüfskript**

- `pruefen_einlesen`: Inventarvorlage (Art, Methode, Satz, Datum, Abgang, KOST1-Zuordnung, Fehler), Aufschlüsselung (Beschriftungsabgleich, Zwischensumme, leere Gruppe), Vorrang des Anlagenverzeichnisses.
- `pruefen`, Fälle „Anlagen“: KSt 1 reproduziert die Stammdaten des Testobjekts aus dem Verzeichnis. KSt 2 mit auslaufender Außenanlage (AfA 30.000, 27.000, 24.000), Anlage im Bau, degressiver Anlage und Statusfällen. Dazu ein Verkauf mit Buchwert aus den Anlagen, nur die Aufschlüsselung ohne Verzeichnis und der Stand gleich Basisjahr.

Vorlagen: `vorlagen/Inventar_Vorlage.xlsx` (Format des DATEV-Exports, erfundene Werte), `vorlagen/Kostenstellen_BWA_Vorlage.xlsx` mit Aufschlüsselung unter der BWA.

**Offen**

- [ ] Grund und Boden ohne KOST1 zuordnen (im Muster 22 von 32 G+B-Anlagen), dann ist AK G+B keine Annahme mehr.
- [ ] Anlagen im Bau: Fertigstellung und AfA-Beginn erfassen (Art Gebäude, Methode linear, Satz).
- [ ] BGA einer Kostenstelle beim Verkauf: geht sie mit ab oder bleibt sie? Derzeit geht sie mit.

