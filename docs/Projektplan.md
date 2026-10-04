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

**Nicht im MVP:** Zins, Tilgung, Restschuld, Vorfälligkeit; Sensitivitäten; IRR; Steuer auf Ausschüttungsebene; Steuerwelten Privat/GbR und gewerbliche Personengesellschaft (siehe Abschnitt 18). Diese Punkte sind in der Situationsbeschreibung als spätere Stufe vermerkt.

## 2. Blattarchitektur

Der Datenfluss läuft in eine Richtung: Eingabeblätter speisen die Rechenblätter, diese die Auswertung. Rückbezüge gibt es nur innerhalb der Rechnung, etwa wenn eine Reinvestition die Prognose eines Neuobjekts anstößt.

```text
Eingabe:  Parameter, Objekte, Verkäufe, Neuobjekte
            ↓
Rechnung: Prognose ⇄ Rücklagen → Liquidität
            ↓
Ausgabe:  Auswertung (inkl. Szenariovergleich)
```

So bleibt nachvollziehbar, woher jede Zahl kommt: Alle Eingaben links, die Rechnung in der Mitte, die Ergebnisse rechts.

## 3. Blätter im Detail

Acht Blätter, getrennt nach Eingabe, Rechnung und Ausgabe. Eingabeblätter sind die einzige Stelle, an der getippt wird.

| Blatt | Typ | Schlüsselfelder | Zweck |
| --- | --- | --- | --- |
| Parameter | Eingabe | Steuerwelt-Schalter (GmbH, Privat/GbR vermögensverwaltend, gewerblich), Grenzsteuersatz, Mietsteigerung, Erhaltungssteigerung, Wertsteigerung, GrESt-Satz, aktives Szenario | alle globalen Annahmen und Schalter |
| Objekte | Eingabe | ObjektID, AK Gebäude, AK G+B, AfA-Satz, Kaufjahr, Restbuchwert 2026, Miete 2026, Erhaltung 2026 | Stammdaten je Bestandsobjekt |
| Verkäufe | Eingabe | ObjektID, Verkaufsjahr, Verkaufspreis oder Faktor, Verkaufskosten, 6b-Nutzung (ja/nein) | ein Datensatz je geplantem Verkauf |
| Neuobjekte | Eingabe | NeuID, Kaufjahr, Kaufpreis, Anteil G+B, AfA-Methode, Mietrendite, Erhaltungsquote, Quelle-Rücklage | Reinvestitionsobjekte |
| Prognose | Rechnung | ObjektID × Jahr (2027-2046), Miete, Erhaltung, AfA, Buchwert, Ergebnis, aktiv-Flag | Jahresmatrix je Objekt, Herzstück |
| Rücklagen | Rechnung | RücklageID, Verkaufsjahr, Betrag G+B, Betrag Gebäude, Fristjahr, Übertrag, Auflösung, Zuschlag | § 6b-Spiegel je Rücklage |
| Liquidität | Rechnung | Jahr, Buchwert-Rückfluss, Steuer, Eigenkapital Reinvest, Alternativanlage | Geldfluss je Jahr |
| Auswertung | Ausgabe | Jahr, Gesamt-GuV, Steuer, stille Reserven, Szenariovergleich | Kennzahlen und Vergleich |

Konvention: ObjektID ist der Schlüssel, der Objekte, Verkäufe, Rücklagen und Prognose verbindet. Neuobjekte bekommen eine eigene ID, laufen in der Prognose aber in derselben Matrix.

## 4. Rechenkern in Formeln

Vier Rechenbausteine, alle als Zellformeln. Die Notation unten ist fachlich, in Excel werden daraus SUMMEWENNS-, WENN- und MAX-Formeln über die Jahresspalten.

**AfA-Fortschreibung je Objekt und Jahr t**

- Miete\_t = Miete\_2026 × (1 + Mietsteigerung)^(t − 2026)
- Erhaltung\_t = Erhaltung\_2026 × (1 + Erhaltungssteigerung)^(t − 2026)
- AfA\_t = MIN(AK\_Gebäude × AfA-Satz; Restbuchwert\_(t−1)), also keine AfA mehr, wenn der Buchwert null ist
- Buchwert\_t = Restbuchwert\_(t−1) − AfA\_t
- Ergebnis\_t = (Miete\_t − Erhaltung\_t − AfA\_t) × aktiv-Flag
- aktiv-Flag = 1, solange t ≤ Verkaufsjahr, sonst 0

**Verkaufsaufteilung im Verkaufsjahr**

- Restbuchwert gesamt = Buchwert Gebäude + AK G+B
- Veräußerungsgewinn = Verkaufspreis − Verkaufskosten − Restbuchwert gesamt
- Gewinn Gebäude = Erlösanteil Gebäude − Buchwert Gebäude; Gewinn G+B = Erlösanteil G+B − AK G+B
- Erlösanteile werden im gleichen Verhältnis wie die Buchwerte oder per Verkehrswert aufgeteilt (Parameter)

**Rücklagenspiegel § 6b**

- Bei 6b-Nutzung = ja und Vorbesitzzeit ≥ 6 Jahre: Rücklage = Veräußerungsgewinn, getrennt nach G+B und Gebäude
- Fristjahr = Verkaufsjahr + 4
- Übertrag im Kaufjahr eines Neuobjekts: Gebäudegewinn nur auf Gebäudeanteil, G+B-Gewinn auf beides
- Neue AfA-Basis Gebäude = Gebäude-AK Neuobjekt − übertragener Gebäudegewinn
- Nicht genutzt bis Fristjahr: Auflösung + 6 % × Jahre als Ertrag

**Szenariovergleich**

- Szenario A (6b-Kette): Steuer im Verkaufsjahr = 0 bei voller Rücklage; dafür spätere AfA geringer
- Szenario B (sofort versteuern): Steuer = Veräußerungsgewinn × Grenzsteuersatz; freigesetztes Kapital in Alternativanlage mit eigener Rendite
- Vergleich über Endvermögen und kumulierte Steuer nach 20 Jahren

## 5. VBA-Module

VBA steuert nur, es rechnet nicht. Die Makros schreiben Eingabewerte und lösen Neuberechnung aus; die Ergebnisse entstehen in den Formeln.

| Modul | Aufgabe |
| --- | --- |
| modObjekte | Objekt anlegen, duplizieren, ausblenden; Prognosezeilen je Objekt erzeugen |
| modSzenario | aktuelles Szenario speichern, laden und benennen; A gegen B vergleichen |
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

Prüfungen schreiben ihr Ergebnis in eine Statusspalte, nicht in Pop-ups allein, damit Fehler im Blatt sichtbar bleiben.

## 6. Umsetzungsreihenfolge

In Etappen, jede mit prüfbarem Zwischenstand. Erst wenn eine Etappe an einem Objekt stimmt, kommt die nächste.

1. **Gerüst:** Parameter- und Objektblatt anlegen, ein Testobjekt erfassen. Prüfbar: Stammdaten vollständig.
2. **AfA-Fortschreibung:** Prognosematrix für ein Objekt über 20 Jahre, ohne Verkauf. Prüfbar: Buchwert läuft korrekt auf null, AfA stoppt danach.
3. **Indexierung:** Miete und Erhaltung mit Steigerungsraten. Prüfbar: Werte wachsen wie erwartet.
4. **Verkauf:** Verkaufsblatt, Aufteilung in Buchwert und Gewinn, Objekt ab Folgejahr inaktiv. Prüfbar: Gewinn = Preis minus Buchwert minus Kosten.
5. **Rücklage:** Rücklagenspiegel, Bildung und Fristjahr. Prüfbar: Gewinn landet getrennt nach G+B und Gebäude in der Rücklage.
6. **Reinvestition:** Neuobjekt, Übertrag, geminderte AfA-Basis. Prüfbar: neue AfA-Basis stimmt, Gebäudeanteil reicht.
7. **Liquidität und Auswertung:** Geldfluss und Gesamt-GuV. Prüfbar: Summen über alle Objekte.
8. **Szenariovergleich:** A gegen B über 20 Jahre. Prüfbar: beide Pfade nachvollziehbar.
9. **VBA-Steuerung und Prüfungen:** erst wenn die Formeln stehen. Prüfbar: Objekt anlegen und Szenario wechseln ohne Formelbruch.

Erst nach Etappe 9 folgt die zweite Stufe mit der Finanzierung.

## 7. Testfälle und Abnahme

Jeder Testfall ist eine Rechnung von Hand, gegen die das Blatt geprüft wird.

| Fall | Eingabe | Erwartetes Ergebnis |
| --- | --- | --- |
| AfA-Ende | Gebäude 800.000, 2,5 %, Kauf 2007, Restbuchwert 2026 400.000 | AfA 20.000 je Jahr, Buchwert Ende 2046 null |
| Verkauf mit Gewinn | Preis 1,4 Mio, Buchwert 680.000, Kosten 0 | Gewinn 720.000 |
| Rücklage voll | 6b ja, Gewinn 720.000 | Steuer im Verkaufsjahr 0, Rücklage 720.000 |
| Übertrag | Neuobjekt Gebäude-AK 900.000 | AfA-Basis 900.000 minus Gebäudegewinn |
| Frist verpasst | keine Reinvestition bis Fristjahr | Auflösung + 6 % je Jahr |
| Szenariovergleich | A gegen B, gleiche Objekte | zwei Endvermögen, Differenz nachvollziehbar |

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
| E | Erhaltung | Formel |
| F | AfA | Formel |
| G | Buchwert Gebäude Ende | Formel |
| H | Ergebnis vor Finanzierung | Formel |

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

**Umsetzung (Stand Etappe 2):** Der Generator legt je Zeile des Objektblatts einen festen Block von 20 Jahreszeilen an, also 200 × 20 = 4.000 Zeilen. Daraus ergeben sich zwei Vereinfachungen gegenüber den Beispielformeln oben:

- Die Zeilen verweisen direkt auf ihre Objektzeile (`Objekte!$J$2` usw.) statt per SVERWEIS über die ObjektID. Das ist schneller und eindeutig, auch bei doppelten IDs.
- Buchwert\_Vorjahr ist die Spalte G der Zeile darüber, im ersten Jahr der Restbuchwert. SUMMEWENNS ist dafür nicht nötig.

Das aktiv-Flag ist vorerst 1, wenn der Status der Objektzeile „OK“ lautet; Objekte mit Fehlern rechnen also nicht mit. Etappe 4 ergänzt das Verkaufsjahr. Neuobjekte (Etappe 6) bekommen eigene Blöcke in derselben Matrix. Miete und Erhaltung werden bereits wie oben indexiert; Etappe 3 prüft das gezielt. Die Spalten sind als `prg_ID`, `prg_Jahr`, `prg_Aktiv`, `prg_Miete`, `prg_Erhaltung`, `prg_AfA`, `prg_Buchwert` und `prg_Ergebnis` benannt; das Blatt ist ohne Kennwort geschützt, Filtern bleibt möglich.

**Abnahme Etappe 2:** Für ein Objekt mit Gebäude 800.000, AfA 2,5 % (40 Jahre), Kauf 2007 muss der Buchwert Ende 2046 null erreichen; die AfA beträgt bis dahin 20.000 je Jahr. Miete und Erhaltung laufen unabhängig davon weiter. Den AfA-Stopp bei Buchwert null prüft zusätzlich ein Fall mit kleinerem Restbuchwert, dessen Ende innerhalb des Rasters liegt.

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
3. Zielzellen auslesen, etwa Buchwert 2046 oder Veräußerungsgewinn.
4. Mit der erwarteten Zahl vergleichen, auf den Euro genau, Toleranz ein Cent für Rundung.
5. Bei Abweichung: Fall, Zelle, Soll und Ist ausgeben und das Skript mit Fehlercode beenden.

**Nutzen**

So wird jede Änderung an den Formel-Bausteinen automatisch gegen alle sechs Fälle geprüft, bevor die Mappe an den Mandanten geht. Das Skript ist das Sicherheitsnetz zwischen Formeländerung und Auslieferung.

## 11. Verkauf (Etappe 4)

Im Verkaufsjahr wird der Erlös in den steuerneutralen Buchwert-Rückfluss und den steuerpflichtigen Gewinn zerlegt, getrennt nach Gebäude und Grund und Boden. Ab dem Folgejahr ist das Objekt inaktiv; das aktiv-Flag aus Etappe 2 schaltet es ab.

**Spaltenlayout Blatt Verkäufe**

| Spalte | Feld | Quelle |
| --- | --- | --- |
| A | ObjektID | aus Objekte |
| B | Verkaufsjahr | Eingabe |
| C | Verkaufspreis | Eingabe, oder Faktor × Jahresmiete |
| D | Verkaufskosten | Eingabe |
| E | Aufteilungsmethode | Dropdown: Buchwert oder Verkehrswert |
| F | 6b-Nutzung | Dropdown: ja oder nein |
| G | Buchwert Gebäude im Verkaufsjahr | Formel, aus Prognose |
| H | Erlösanteil Gebäude | Formel |
| I | Gewinn Gebäude | Formel |
| J | Gewinn G+B | Formel |
| K | Veräußerungsgewinn gesamt | Formel |

**Formeln (Beispiel Zeile 2)**

Den Buchwert des Gebäudes im Verkaufsjahr holt sich die Zeile aus der Prognosematrix, über ObjektID und Verkaufsjahr.

```text
G2  =SUMMEWENNS(Prognose!G:G; Prognose!A:A; A2; Prognose!B:B; B2)
```

Der Erlös wird aufgeteilt. Bei Methode Buchwert nach dem Verhältnis der Buchwerte, bei Verkehrswert nach einem separat gepflegten Anteil.

```text
H2  =WENN(E2="Buchwert"; (C2-D2) * G2/(G2+SVERWEIS(A2;obj_Basis;SPALTE_AKGuB;FALSCH)); (C2-D2)*SVERWEIS(A2;obj_Basis;SPALTE_VKQuoteGeb;FALSCH))
```

Der Gewinn je Teil ist der jeweilige Erlösanteil minus Buchwert. Beim Grund und Boden ist der Buchwert die ursprüngliche Anschaffung, da darauf keine AfA läuft.

```text
I2  =H2 - G2
J2  =((C2-D2) - H2) - SVERWEIS(A2;obj_Basis;SPALTE_AKGuB;FALSCH)
K2  =I2 + J2
```

**Verknüpfung zur Prognose:** Das aktiv-Flag in Spalte C der Prognose greift bereits auf Verkaeufe Spalte B zu. Damit endet das laufende Ergebnis des Objekts automatisch im Jahr nach dem Verkauf. Der Veräußerungsgewinn K fließt in Etappe 5 in den Rücklagenspiegel oder, bei 6b-Nutzung nein, direkt in die Steuer.

**Abnahme Etappe 4:** Preis 1,4 Mio, Verkaufskosten 0, Buchwert gesamt 680.000 muss einen Veräußerungsgewinn von 720.000 ergeben. Bei hälftiger Aufteilung liegt der Gebäudeanteil des Gewinns korrekt getrennt vom G+B-Anteil vor.

## 12. Rücklagenspiegel § 6b (Etappe 5)

Bei 6b-Nutzung ja wird der Veräußerungsgewinn in eine Rücklage eingestellt, getrennt nach Gebäude und Grund und Boden, weil die Übertragbarkeit unterschiedlich ist. Die Rücklage stundet die Steuer bis zur Reinvestition oder bis zum Fristablauf.

**Spaltenlayout Blatt Rücklagen**

| Spalte | Feld | Quelle |
| --- | --- | --- |
| A | RücklageID | je Verkauf mit 6b ja |
| B | ObjektID Herkunft | aus Verkäufe |
| C | Verkaufsjahr | aus Verkäufe |
| D | Vorbesitzzeit Jahre | Formel, Verkaufsjahr minus Kaufjahr |
| E | Betrag Gebäude | aus Verkäufe, Gewinn Gebäude |
| F | Betrag G+B | aus Verkäufe, Gewinn G+B |
| G | Fristjahr | Formel, Verkaufsjahr + 4 |
| H | übertragen Gebäude | aus Neuobjekte |
| I | übertragen G+B | aus Neuobjekte |
| J | Restrücklage | Formel |
| K | Auflösung im Fristjahr | Formel |
| L | Gewinnzuschlag | Formel |

**Formeln (Beispiel Zeile 2)**

Vorbesitzzeit und Fristjahr sind einfache Differenzen. Die Vorbesitzzeit muss mindestens sechs Jahre sein, sonst ist die Rücklage unzulässig.

```text
D2  =C2 - SVERWEIS(B2;obj_Basis;SPALTE_Kaufjahr;FALSCH)
G2  =C2 + 4
```

Die Restrücklage ist der eingestellte Betrag minus das, was bereits auf Neuobjekte übertragen wurde.

```text
J2  =(E2 + F2) - (H2 + I2)
```

Ist im Fristjahr noch eine Restrücklage offen, wird sie aufgelöst und mit sechs Prozent je vollem Jahr verzinst. Vor dem Fristjahr sind Auflösung und Zuschlag null.

```text
K2  =WENN(aktuelles_Jahr>=G2; J2; 0)
L2  =WENN(aktuelles_Jahr>=G2; J2 * 0,06 * 4; 0)
```

**Wichtige Regel zur Übertragbarkeit:** Der Gebäudegewinn in Spalte E darf nur auf den Gebäudeanteil eines Neuobjekts übertragen werden, der G+B-Gewinn in Spalte F auf Gebäude oder Grund und Boden. Diese Trennung prüft modPruefung in Etappe 9.

**Abnahme Etappe 5:** Veräußerungsgewinn 720.000 bei 6b ja und Vorbesitzzeit mindestens sechs Jahre muss im Verkaufsjahr zu Steuer null führen; die Rücklage steht mit 720.000, getrennt in Gebäude- und G+B-Anteil.

## 13. Reinvestition (Etappe 6)

Ein Neuobjekt nimmt die Rücklage auf. Der übertragene Gewinn mindert die AfA-Basis des neuen Gebäudes, dadurch läuft die AfA künftig von einem niedrigeren Wert. Das ist der Kern der Stundung: keine Steuer heute, dafür weniger Abschreibung morgen.

**Spaltenlayout Blatt Neuobjekte**

| Spalte | Feld | Quelle |
| --- | --- | --- |
| A | NeuID | eigene ID |
| B | Kaufjahr | Eingabe |
| C | Kaufpreis | Eingabe |
| D | Anteil G+B | Eingabe, Prozent |
| E | Kaufnebenkosten | Eingabe oder Formel |
| F | AfA-Methode | Dropdown: linear 3 %, degressiv 5 %, Bestand 2 % |
| G | Quelle RücklageID | Dropdown aus Rücklagen |
| H | AK Gebäude brutto | Formel |
| I | übertragener Gebäudegewinn | Formel |
| J | AfA-Basis Gebäude | Formel |

**Formeln (Beispiel Zeile 2)**

Die Gebäude-Anschaffung ist der Kaufpreis ohne G+B-Anteil, plus die auf das Gebäude entfallenden Nebenkosten. Nebenkosten wie Grunderwerbsteuer und Notar werden aktiviert, nicht sofort abgezogen.

```text
H2  =(C2 * (1 - D2) ) + (E2 * (1 - D2))
```

Der übertragbare Gewinn kommt aus dem Rücklagenspiegel. Übertragen werden darf höchstens der Gebäudeanteil des Neuobjekts.

```text
I2  =MIN( SVERWEIS(G2;Ruecklagen!A:E;5;FALSCH); H2 )
J2  =H2 - I2
```

Die neue AfA-Basis J fließt zurück in die Objektlogik: Das Neuobjekt bekommt eine Zeile in der Prognosematrix wie ein Bestandsobjekt, nur dass seine Gebäude-AK der geminderte Wert aus Spalte J ist und die AfA im Kaufjahr beginnt.

**Rückkopplung in den Rücklagenspiegel:** Der übertragene Betrag I landet in Spalte H des Rücklagenspiegels und senkt dort die Restrücklage. Ist die Rücklage voll übertragen, entfällt die spätere Auflösung samt Zuschlag.

**Abnahme Etappe 6:** Neuobjekt mit Gebäude-AK 900.000 und übertragenem Gebäudegewinn 720.000 muss eine AfA-Basis von 180.000 ergeben. Die AfA des Neuobjekts läuft ab Kaufjahr von diesen 180.000, nicht von 900.000.

## 14. Liquidität und Auswertung (Etappe 7)

Die Prognosematrix liefert je Objekt und Jahr die Einzelwerte. Liquidität und Auswertung fassen sie über alle Objekte zu Jahreswerten zusammen, per SUMMEWENNS über das Jahr. Alle Angaben sind vor Finanzierung.

**Blatt Liquidität, je Jahr eine Zeile**

| Spalte | Feld | Formel-Idee |
| --- | --- | --- |
| A | Jahr | 2027 bis 2046 |
| B | laufendes Ergebnis | SUMMEWENNS über Prognose Ergebnis, Jahr = A |
| C | Buchwert-Rückfluss aus Verkauf | SUMMEWENNS über Verkäufe, steuerneutraler Teil |
| D | Veräußerungsgewinn | SUMMEWENNS über Verkäufe, Jahr = A |
| E | Steuer | abhängig von 6b: 0 bei Rücklage, sonst Gewinn × Satz |
| F | Eigenkapital in Reinvestition | aus Neuobjekte, Kaufjahr = A |
| G | freier Mittelzufluss | B + C + D − E − F |

**Blatt Auswertung, Kennzahlen je Jahr**

| Spalte | Feld | Formel-Idee |
| --- | --- | --- |
| A | Jahr | 2027 bis 2046 |
| B | Gesamt-GuV vor Finanzierung | laufendes Ergebnis plus steuerpflichtiger Gewinn |
| C | Steuer | aus Liquidität Spalte E |
| D | Ergebnis nach Steuer | B − C |
| E | stille Reserven | Summe aus Verkehrswert minus Buchwert je aktivem Objekt |
| F | kumulierte Steuer | laufende Summe über C |

**Zur Steuerformel in Spalte E der Liquidität**

Die Steuer hängt am 6b-Schalter des jeweiligen Verkaufs. Bei Rücklage ist sie im Verkaufsjahr null, der Gewinn ist gestundet. Ohne Rücklage fällt sie sofort an, mit dem Grenzsteuersatz vom Parameterblatt.

```text
E2  =SUMMENPRODUKT((Verkaeufe!$B$2:$B$999=A2)*(Verkaeufe!$F$2:$F$999="nein")*Verkaeufe!$K$2:$K$999) * par_Steuersatz
```

**Stille Reserven als Kennzahl**

Stille Reserven zeigen, wie viel unversteuerter Wert im Bestand steckt: der Verkehrswert minus Buchwert über alle noch aktiven Objekte. Das ist keine Steuerposition, sondern eine Steuerungsgröße für die Entscheidung, wann sich ein Verkauf lohnt.

**Abnahme Etappe 7:** Die Summe des laufenden Ergebnisses über alle Objekte eines Jahres muss mit der Einzelsumme aus der Prognosematrix übereinstimmen. In einem Verkaufsjahr ohne 6b muss die Steuer gleich Veräußerungsgewinn mal Satz sein.

## 15. Szenariovergleich (Etappe 8)

Der Vergleich beantwortet die Kernfrage des Mandanten: Lohnt die § 6b-Kette, oder ist es besser, die Steuer sofort zu zahlen und das freie Kapital anderweitig anzulegen? Beide Pfade laufen über dieselben Objekte und 20 Jahre, nur die Behandlung des Veräußerungsgewinns unterscheidet sich.

**Die zwei Szenarien**

- **Szenario A, 6b-Kette:** Gewinn in die Rücklage, keine Steuer im Verkaufsjahr, Reinvestition in ein Neuobjekt mit geminderter AfA-Basis. Das Neuobjekt wirft Miete ab, die AfA ist aber kleiner.
- **Szenario B, sofort versteuern:** Gewinn wird im Verkaufsjahr versteuert. Das verbleibende Kapital geht in eine Alternativanlage mit eigener Rendite, statt in eine Immobilie.

**Umsetzung im Blatt**

Ein Schalter auf dem Parameterblatt, par\_Szenario, steuert, welcher Pfad gerechnet wird. Für den Vergleich werden beide Läufe gespeichert: VBA rechnet A, kopiert die Endwerte in eine Vergleichsspalte, rechnet B, kopiert ebenso. So stehen beide Ergebnisse nebeneinander, ohne zwei komplette Mappen.

**Vergleichskennzahlen nach 20 Jahren**

| Kennzahl | Szenario A | Szenario B |
| --- | --- | --- |
| Endvermögen Immobilien plus Anlage | aus Auswertung | aus Auswertung |
| kumulierte Steuer | geringer, aber später | höher, aber sofort |
| laufende Mieterträge | inklusive Neuobjekt | nur Alt plus Alternativanlage |
| verlorene AfA durch Minderung | ja | nein |

**Die ehrliche Kennzahl**

Entscheidend ist das Endvermögen nach Steuern, nicht die gesparte Steuer allein. Szenario A spart Steuer heute, verliert aber AfA und bindet Kapital in Immobilien. Ob sich das lohnt, hängt an der Rendite der Alternativanlage und an der Wertsteigerung der Neuimmobilie. Genau diesen Vergleich macht das Blatt sichtbar.

**Abnahme Etappe 8:** Beide Pfade liefern je ein Endvermögen; die Differenz ist nachvollziehbar aus gestundeter Steuer, verlorener AfA und Alternativrendite. Bei Alternativrendite null und gleicher Wertentwicklung muss A vorn liegen, weil die Steuerstundung dann reiner Zinsvorteil ist.

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

**Austauschbare Einleseschicht**

Wie bei der Pillar-2-Pipeline kapselt ein eigenes Modul das Einlesen. Ändert sich das Quellformat, wird nur diese Schicht angepasst, nicht der Rest.

**Offene Punkte, nächste Woche in der Arbeit zu prüfen**

- [ ] Gibt es ein Anlageverzeichnis mit Anschaffungskosten und Buchwerten je Objekt?
- [ ] Ist darin die Aufteilung Gebäude zu Grund und Boden schon enthalten?
- [ ] Haben die Kostenstellenblätter ein einheitliches Layout mit fester ObjektID?

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

## 18. Steuerwelt-Schalter

Der Schalter par\_Steuerwelt auf dem Parameterblatt legt die Rechtsform fest. Sie bestimmt Steuersatz, Behandlung des Veräußerungsgewinns und ob § 6b überhaupt möglich ist.

| | GmbH | Privat / GbR vermögensverwaltend | gewerblich (Personengesellschaft) |
| --- | --- | --- | --- |
| Steuersatz | KSt + SolZ, GewSt ggf. durch erweiterte Kürzung nahe null | persönlicher Grenzsteuersatz | persönlicher Satz + GewSt (teilweise angerechnet) |
| Veräußerungsgewinn | immer steuerpflichtig | steuerfrei nach 10 Jahren Haltedauer (§ 23 EStG) | immer steuerpflichtig |
| § 6b-Rücklage | ja | nein, nur im Betriebsvermögen | ja |
| Wirtschaftsgebäude-AfA 3 % | ja | nein | ja |

**MVP:** Gerechnet wird nur die GmbH. Bei jeder anderen Auswahl zeigt das Parameterblatt den Status "nicht im MVP", damit niemand unbemerkt mit GmbH-Logik für eine andere Rechtsform rechnet.

**Spätere Etappe Privat/GbR:** Szenario A (§ 6b-Kette) entfällt; an seine Stelle tritt die 10-Jahres-Regel. Der Szenariovergleich wird zu "Verkauf vor oder nach Ablauf der Haltefrist". Bei der GbR entscheidet die Einordnung (vermögensverwaltend oder gewerblich, etwa durch gewerblichen Grundstückshandel), welche Spalte gilt; die Warnung zur Drei-Objekt-Grenze aus modPruefung liefert dafür den Hinweis.

Einordnung und Sätze vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
