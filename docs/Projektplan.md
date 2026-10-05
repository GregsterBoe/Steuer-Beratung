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
| Parameter | Eingabe | Steuerwelt-Schalter, Grenzsteuersatz, Mietsteigerung, Erhaltungssteigerung, Kostensteigerung, Wertsteigerung, GrESt-Satz, aktives Szenario | alle globalen Annahmen und Schalter |
| Objekte | Eingabe | ObjektID, AK Gebäude, AK G+B, AfA-Satz, Kaufjahr, Restbuchwert 2026, Miete 2026, Erhaltung 2026 | Stammdaten je Bestandsobjekt |
| Verkäufe | Eingabe | ObjektID, Verkaufsjahr, Verkaufspreis oder Faktor, Verkaufskosten, 6b-Nutzung (ja/nein) | ein Datensatz je geplantem Verkauf |
| Neuobjekte | Eingabe | NeuID, Kaufjahr, Kaufpreis, Anteil G+B, AfA-Methode, Mietrendite, Erhaltungsquote, Quelle-Rücklage | Reinvestitionsobjekte |
| Prognose | Rechnung | ObjektID × Jahr (2027-2046), Miete, Erhaltung, AfA, Buchwert, Ergebnis, aktiv-Flag | Jahresmatrix je Objekt, Herzstück |
| Rücklagen | Rechnung | RücklageID, Verkaufsjahr, Betrag G+B, Betrag Gebäude, Fristjahr, Übertrag, Auflösung, Zuschlag | § 6b-Spiegel je Rücklage |
| Liquidität | Rechnung | Jahr, Buchwert-Rückfluss, Steuer, Eigenkapital Reinvest, Alternativanlage | Geldfluss je Jahr |
| Auswertung | Ausgabe | Jahr, Gesamt-GuV, Steuer, stille Reserven, Szenariovergleich | Kennzahlen und Vergleich |
| Übersicht | Ausgabe | Jahr, Verkehrswert Baseline, Verkehrswert Plan, Differenz, Diagramm | Gesamtwert des Bestands im Jahresverlauf, Plan gegen Nichtstun |

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
- Fristjahr = Verkaufsjahr + 4
- Übertrag im Kaufjahr eines Neuobjekts: Gebäudegewinn nur auf Gebäudeanteil, G+B-Gewinn auf beides (§ 6b Abs. 1 EStG)
- Reihenfolge: Gebäudegewinn auf neues Gebäude, G+B-Gewinn zuerst auf neuen G+B (kostet keine AfA), nur der Rest auf das Gebäude (Abschnitt 13)
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
| AfA-Ende | Gebäude 800.000, 2,5 %, Kauf 2007, Restbuchwert 2026 400.000 | Buchwert Ende 2046 null, AfA danach 0 |
| AfA läuft im Raster aus | Gebäude 800.000, 2 %, Restbuchwert 2026 50.000 | AfA 2030 nur 2.000, ab 2031 AfA und Buchwert 0, Miete läuft weiter |
| Verkauf mit Gewinn | Preis 1,4 Mio, Buchwert 680.000, Kosten 0, hälftig | Gewinn 720.000, davon Gebäude 220.000, G+B 500.000 |
| Rücklage voll | 6b ja, Gewinn 720.000 | Steuer im Verkaufsjahr 0, Rücklage 720.000 |
| Verkauf mit Kaufvertragsaufteilung | Preis 1,4 Mio, Kosten 40.000, 30 % G+B, Verkauf Ende 2030 | Gewinn Gebäude 536.000, G+B 208.000 |
| Übertrag | Rücklage Gebäude 220.000, G+B 500.000; Neuobjekt G+B 360.000, Gebäude 840.000 | AfA-Basis 480.000, AK G+B 0 (G+B-Gewinn zuerst auf G+B) |
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
| E | weitere Einnahmen | Formel |
| F | Erhaltung | Formel |
| G | weitere Ausgaben | Formel |
| H | AfA | Formel |
| I | Buchwert Gebäude Ende | Formel |
| J | Ergebnis vor Finanzierung | Formel |

Zusätzlich für die Übersicht: K Verkehrswert Ende (prg\_Verkehrswert, Verkehrswert aktuell × (1 + par\_Wertsteig)^(Jahr − Basisjahr), unabhängig vom Verkauf) und L im Bestand Ende (prg\_Bestand, 1 solange Jahr < Verkaufsjahr).

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
| F | § 6b nutzen | vk\_6b | Dropdown ja/nein, wirkt ab Etappe 5 |
| G | Vorbesitzzeit Jahre | vk\_Vorbesitz | Verkaufsjahr − Kaufjahr |
| H | Buchwert Gebäude Ende Verkaufsjahr | vk\_BuchwertGeb | SUMIFS über prg\_Buchwert |
| I | AK G+B | vk\_AKGuB | aus Objekte |
| J | Nettoerlös | vk\_Nettoerloes | C − D |
| K | Anteil G+B verwendet | vk\_AnteilGuB | E, sonst 1 − Verkehrswertanteil Gebäude |
| L | Erlösanteil Gebäude | vk\_ErloesGeb | J − M |
| M | Erlösanteil G+B | vk\_ErloesGuB | J × K |
| N | Gewinn Gebäude | vk\_GewinnGeb | L − H |
| O | Gewinn G+B | vk\_GewinnGuB | M − I, G+B wird nicht abgeschrieben |
| P | Veräußerungsgewinn | vk\_Gewinn | N + O |
| Q | Status | vk\_Status | Plausibilität |

**Formeln (Zeile 2, englische Syntax wie in der Mappe)**

```text
H2  =SUMIFS(prg_Buchwert, prg_ID, A2, prg_Jahr, B2)
K2  =IF(E2<>"", E2, IF(INDEX(obj_VKQuoteGeb, MATCH(A2,obj_ID,0))="", "", 1-INDEX(obj_VKQuoteGeb, MATCH(A2,obj_ID,0))))
M2  =J2*K2
N2  =L2-H2
O2  =M2-I2
```

Die berechneten Spalten bleiben leer, solange die ObjektID fehlt oder unbekannt ist. Der Buchwert in H ist der Prognosewert am Ende des Verkaufsjahrs, die AfA des Verkaufsjahrs ist also schon abgezogen.

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

**Verknüpfung zur Prognose:** Das aktiv-Flag liest das Verkaufsjahr per INDEX/MATCH aus vk\_Jahr; prg\_Bestand ist schon im Verkaufsjahr 0. Der Veräußerungsgewinn P fließt in Etappe 5 in den Rücklagenspiegel oder, bei § 6b nein, in Etappe 7 direkt in die Steuer. Noch nicht umgesetzt ist ein Verkaufspreis als Faktor × Jahresmiete.

**Abnahme Etappe 4 (geprüft in `pruefung/pruefen.py`):**
- Preis 1,4 Mio, Kosten 0, Buchwert gesamt 680.000 (Gebäude 480.000 Ende 2027, G+B 200.000), Verkehrswertanteil Gebäude 50 %:
  - Veräußerungsgewinn 720.000
  - davon Gebäude 220.000 und G+B 500.000
- Kaufvertrag 30 % G+B, Kosten 40.000, Verkauf Ende 2030:
  - Gewinn Gebäude 536.000
  - Gewinn G+B 208.000

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

**Abnahme Etappe 5:** Veräußerungsgewinn 720.000 bei 6b ja und Vorbesitzzeit mindestens sechs Jahre muss im Verkaufsjahr zu Steuer null führen; die Rücklage steht mit 720.000, getrennt in Gebäude 220.000 und G+B 500.000. Nur positive Teilgewinne gehen in die Rücklage.

**Vor Umsetzung klären:** Frist sechs statt vier Jahre, wenn mit dem Bau eines neuen Gebäudes vor Ende des vierten Jahres begonnen wurde (§ 6b Abs. 3). Ob die AK des neuen G+B bis auf null gekürzt werden dürfen.

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

**Übertragung in fester Reihenfolge (korrigiert).** Der erste Entwurf übertrug nur den Gebäudegewinn und nur auf das Gebäude (`I2 = MIN(Rücklage Spalte E; H2)`). Der G+B-Gewinn ging dabei verloren. Richtig nach § 6b Abs. 1 EStG ist: Ein Gebäudegewinn darf nur auf ein Gebäude übertragen werden, ein G+B-Gewinn auf Gebäude oder auf G+B. Steuerlich günstig ist diese Reihenfolge:

1. **ü1 = MIN(Rücklage Gebäude; AK Gebäude neu).** Der Gebäudegewinn geht auf das neue Gebäude, denn er hat keine andere Verwendung.
2. **ü2 = MIN(Rücklage G+B; AK G+B neu).** Der G+B-Gewinn geht zuerst auf den neuen G+B. Das kostet keine AfA, die Steuer bleibt bis zum Verkauf des Grundstücks gestundet.
3. **ü3 = MIN(Rücklage G+B − ü2; AK Gebäude neu − ü1).** Erst der Rest des G+B-Gewinns geht auf das Gebäude und mindert dessen AfA-Basis.

Danach gilt:
- AfA-Basis Gebäude = AK Gebäude neu − ü1 − ü3
- steuerliche AK G+B = AK G+B neu − ü2
- Rücklage Gebäude sinkt um ü1, Rücklage G+B um ü2 + ü3. Was übrig bleibt, wartet auf ein weiteres Neuobjekt oder wird im Fristjahr mit Zuschlag aufgelöst.

Beispiel: G+B-Gewinn 4, neuer G+B kostet 3. Dann werden 3 beim G+B abgezogen und 1 beim Gebäude.

Spalten statt I und J oben: I ü1, J ü2, K ü3, L AfA-Basis Gebäude, M steuerliche AK G+B.

Die neue AfA-Basis fließt zurück in die Objektlogik: Das Neuobjekt bekommt Zeilen in der Prognosematrix wie ein Bestandsobjekt, mit der geminderten AfA-Basis und AfA ab dem Kaufjahr. Ein Neuobjekt gehört nur in die Plan-Linie der Übersicht, nicht in die Baseline.

**Rückkopplung in den Rücklagenspiegel:** ü1 landet in Spalte H (übertragen Gebäude), ü2 + ü3 in Spalte I (übertragen G+B). Ist die Rücklage voll übertragen, entfallen Auflösung und Zuschlag.

**Abnahme Etappe 6:**
- Ausgangslage: Rücklage aus dem Abnahmefall der Etappe 4 (Gebäude 220.000, G+B 500.000). Neuobjekt mit Kaufpreis 1,2 Mio, Anteil G+B 30 %, also G+B 360.000 und Gebäude 840.000.
- Übertragung: ü1 = 220.000, ü2 = 360.000, ü3 = 140.000.
- Ergebnis: AfA-Basis Gebäude 480.000, steuerliche AK G+B 0, Restrücklage 0.
- Die AfA des Neuobjekts läuft von 480.000, nicht von 840.000.

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

**Layout eines Kostenstellenblatts (Muster liegt vor)**

Die Zeilen folgen der DATEV-BWA Form 01 (Kurzfristige Erfolgsrechnung). Die Nummern in Spalte B sind BWA-Zeilennummern, keine Sachkonten. Die BWA ist kontenrahmenunabhängig; ob dahinter SKR03 oder SKR04 gebucht wird, geht aus dem Blatt nicht hervor.

| Zelle/Bereich | Inhalt |
| --- | --- |
| B2 | Kostenstelle, z. B. „KSt 1“ |
| C2 | Objektbezeichnung, z. B. „KC 24+26“ |
| Zeile 4 | Kopf: B „Nr.“, C „Bezeichnung kurz“, F Vorjahr (2025), G–R Monate 2026, S Summe 2026, T–AM Jahre 2027–2046 |
| ab Zeile 6 | eine Zeile je BWA-Position, Schlüssel ist die Nummer in Spalte B |
| G–R | Ist-Werte bis zum letzten gebuchten Monat, danach Hochrechnung per Mittelwert |

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

- [ ] Gibt es ein Anlageverzeichnis mit Anschaffungskosten und Buchwerten je Objekt?
- [ ] Ist darin die Aufteilung Gebäude zu Grund und Boden schon enthalten?
- [x] Haben die Kostenstellenblätter ein einheitliches Layout mit fester ObjektID? Muster liegt vor (B2 Kostenstelle, C2 Objekt), Einheitlichkeit über alle Blätter noch bestätigen.
- [ ] Welcher Kontenrahmen (SKR03 oder SKR04) wird gebucht? Für die BWA-Werte egal, relevant erst beim Abgleich mit Sachkonten.

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
