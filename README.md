# Prognosemodell V+V

20-Jahres-Prognose (2027–2046, Basis 2026) für Immobilien im Betriebsvermögen einer GmbH: AfA-Fortschreibung, Verkauf, § 6b-Rücklage und Reinvestition. Am Ende werden Szenarien verglichen: 6b-Kette, sofort versteuern und das Kapital anlegen, oder sofort versteuern und trotzdem reinvestieren.

## Prinzipien

- **Endprodukt:** eine eigenständige Excel-Datei, die der Mandant ohne Python bedient.
- **Rechnung in Zellformeln:** jeder Schritt ist im Blatt nachvollziehbar. VBA steuert nur Objekte, Szenarien, Prüfungen und Neuberechnung.
- **Python als Generator:** openpyxl baut die Mappe einmalig beim Erstellen. Ein Prüfskript rechnet sie per LibreOffice headless durch und gleicht sie mit den Testfällen ab.
- **Alle Annahmen auf dem Parameterblatt**, keine festen Zahlen in Formeln.
- **Selbstdokumentierend:** Spaltenköpfe zeigen Pflichtfelder (rot, *), optionale Eingaben (blau) und berechnete Spalten (grau); ein Kommentar am Kopf erklärt das Feld bzw. die Herleitung. Auf dem Parameterblatt sind die Kernparameter mit ★ markiert, weniger wichtige Abschnitte eingeklappt, und jeder Parameter sagt, was er beeinflusst (Texte in `prognosemodell/erklaerungen.py`).
- **Finanzierung:** Kredite der Neuobjekte werden je Kauf im Blatt Neuobjekte festgelegt (Tilgungsplan im Blatt Darlehen). Bestandsobjekte rechnen mit dem Zinsaufwand der Buchhaltung (BWA 1310); mit Restschuld, Zinssatz und Rate im Blatt Objekte als Tilgungsplan, beim Verkauf aus dem Erlös abgelöst.

## Blätter

| Eingabe | Rechnung | Ausgabe | Kontrolle | BWA |
| --- | --- | --- | --- | --- |
| Parameter, Objekte, Anlagen, Verkäufe, Neuobjekte | Prognose, Rücklagen, Darlehen, Liquidität | Start, Übersicht, Vergleich, Ergebnisbrücke, Auswertung | Prüfung, Varianten | Alle Objekte, Alle Objekte halten, je Kostenstelle ein Blatt, Verkauf und Kauf |

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
- Ausgabe im DATEV-BWA-Format je Kostenstelle mit Summenblatt, Sonderbereich Verkauf und Kauf
- Blatt BWA-Zuordnung: BWA-Zeile je Sonderposten (Verkauf netto oder brutto, § 6b-Rücklage, Neuobjekte, Zins) wählbar, Herleitung unter jeder BWA, Kontrolle gegen die Liquidität
- Startblatt mit Handlungsempfehlung, Auffülllogik für fehlende Daten mit Farblogik, Schnellcheck-Mappe
- Übernahme (`--uebernehmen`): eingetippte Werte einer älteren Mappe gehen in die neu erzeugte über, Spalten nach Überschrift, Objekte und Anlagen nach ID; eingelesene Werte kommen aus der neuen Kostenstellen-Datei
- Blatt Anlagen: Anlagenverzeichnis (DATEV-Export) oder Aufschlüsselung der Abschreibungen aus den Kostenstellenblättern; liefert AK, Buchwert und Kaufjahr je Objekt und die AfA je Anlage und Jahr

## Nutzung

```bash
pip install -r requirements.txt
python -m prognosemodell             # erzeugt ausgabe/Prognosemodell_VV.xlsx (mit Testobjekt)
python -m prognosemodell --ohne-testdaten
python -m prognosemodell --makros    # .xlsm mit VBA-Steuerung (braucht LibreOffice beim Bauen)
python -m prognosemodell --schnellcheck --kostenstellen Kostenstellen.xlsx   # schlanke Mappe ausgabe/Schnellcheck_VV.xlsx
python -m prognosemodell --kostenstellen Kostenstellen.xlsx --ausgabe Ordner/Prognose.xlsx   # laufende Werte je Blatt einlesen; nur Ordner = Standardname darin
python -m prognosemodell --kostenstellen Kostenstellen.xlsx --inventar Inventar_2025.xlsx   # dazu das Anlagenverzeichnis
python -m prognosemodell --kostenstellen Kostenstellen.xlsx --uebernehmen Alte_Mappe.xlsx --ausgabe Neue_Mappe.xlsx   # eingetippte Werte der alten Mappe übernehmen
python -m pruefung.pruefen           # rechnet per LibreOffice headless und prüft gegen Sollwerte
python -m pruefung.pruefen "Etappe 9"  # nur Fälle, deren Name den Text enthält
python -m pruefung.pruefen -j 2      # höchstens 2 Fälle gleichzeitig (Standard: ein Prozess je Kern)
python -m pruefung.pruefen_einlesen  # prüft die Einleseschicht, ohne LibreOffice
python -m pruefung.zirkelbezug      # sucht Zirkelbezüge, wie Excel sie sieht, ohne LibreOffice
python -m prognosemodell.export Prognose.xlsx   # Mandanten-Export Prognose_Mandant.xlsx: kompakt, nur Werte
python -m prognosemodell.vorlagen    # schreibt vorlagen/Kostenstellen_BWA_Vorlage.xlsx und vorlagen/Inventar_Vorlage.xlsx
```

Das Prüfskript und `--makros` brauchen LibreOffice mit Calc und der Python-UNO-Brücke (`soffice`, unter Debian/Ubuntu die Pakete `libreoffice-calc` und `python3-uno`).

Excel meldet einen Zirkelbezug schon, wenn eine Formel einen Bereich nennt, in dem eine von ihr abhängige Zelle liegt, etwa SUMIFS über die ganze Prognosespalte. Das gilt auch, wenn der Wert gar nicht gebraucht wird. LibreOffice rechnet solche Ringe oft trotzdem durch, das Prüfskript sieht sie daher nicht. `pruefung.zirkelbezug` baut den Abhängigkeitsgraphen der erzeugten Mappen und meldet jeden Ring mit einem Pfad zum Nachlesen. Die Verkäufe lesen den Buchwert deshalb nur aus den Zeilen der Bestandsobjekte (`prg_BuchwertObj`), sonst hingen sie über die Rücklage von den Neuobjekten ab.

Das Prüfskript rechnet die Fälle parallel und mit kleinen Mappen (10 Objekte, Verkäufe und Neuobjekte, 5 Neukauf-Kostenstellen statt 200, 50, 50 und 20), denn jede leere Zeile kostet LibreOffice Rechenzeit. Zwei Fälle rechnen zur Kontrolle mit der Mappe in Originalgröße.

Zielformat der Eingabe ist die DATEV-BWA-Kostenstellenblattsammlung; `vorlagen/Kostenstellen_BWA_Vorlage.xlsx` zeigt das Layout mit erfundenen Werten. Als Jahresspalte gilt ein Kopf wie 2026, „Jahr 2026“ oder „Plan 2027“; ein Summenblatt „Alle Objekte“ wird übersprungen. Die Ergebnisse stehen wieder in dieser Struktur, siehe unten.

`--kostenstellen` überspringt jedes Blatt, das nicht im Kostenstellenformat ist (kein „Nr.“ in B4, keine Kostenstelle in B2 oder keine Spalte des Basisjahrs in Zeile 4), etwa Annahmen oder Übersichten. Jedes übersprungene Blatt nennt es mit Grund in der Ausgabe. Je Kostenstellenblatt liest es B2 (Kostenstelle = ObjektID), C2 (Objektname) und aus der Spalte des Basisjahrs die BWA-Zeilen 1020 (Miete), 1090 (weitere Einnahmen), 1250 (Erhaltung) sowie 1100–1220 und 1260 (weitere Ausgaben). Die Datei muss in Excel gespeichert sein, damit berechnete Werte vorliegen. Dazu liest es 1240 als „AfA Basisjahr lt. Buchhaltung“, die schon geplante AfA der Folgejahre aus 1240 (Blatt **AfA-Plan**: Jahre mit Wert ersetzen die Fortschreibung, leere rechnet das Modell) und, falls vorhanden, die Aufschlüsselung der Abschreibungen unter der BWA (Buchwert und Jahres-AfA je Anlagengruppe, siehe unten). Die übrigen steuerlichen Stammdaten (AK, Kaufjahr) kommen aus dem Anlagenverzeichnis; solange sie fehlen, rechnet die Mappe mit Annahmen (siehe unten).

Blätter hinter der Kostenstelle „KSt 9999“ (im Muster KSt 31–35) sind **Neukauf-Kostenstellen**: kein Bestandsobjekt, sondern Planwerte für einen Kauf. Jede bekommt in der Mappe ein eigenes Blatt im BWA-Layout. Das Blatt **Neukauf-KSt** hält die Planwerte je Jahr (Miete, weitere Einnahmen, Erhaltung, weitere Ausgaben): das Basisjahr verweist auf das Blatt der Kostenstelle (dort gelb, Eingabe), die Planjahre sind Eingabe im Blatt Neukauf-KSt, vorbelegt mit den eingelesenen Werten. Jahre ohne Wert schreibt es vom letzten Wert mit der Steigerung fort; eine 0 vor dem ersten anderen Wert (Kostenstelle ohne Buchungen) gilt als leer. Ein Neuobjekt wählt sie in der Spalte „Kostenstelle Neukauf“ (Auswahlliste) und rechnet ab dem Jahr nach dem Kauf mit diesen Werten statt mit Mietrendite und Erhaltungsquote. Das Blatt der Kostenstelle ist dann zugleich das Blatt des Neuobjekts: Die Planjahre kommen grau aus dem Modell, also Miete (1020, aus Planwert oder Mietrendite), weitere Einnahmen, Erhaltung, weitere Ausgaben (nach dem Anteil der Kostenart im Basisjahr, Rest in 1260), AfA (1240) und Kreditzinsen (1310), darunter Kreditauszahlung, Tilgung und Restschuld; ein eigenes Blatt NEU-… entfällt. Über die Prognose zählt das Neuobjekt in „Alle Objekte“. Die AfA rechnet weiter das Modell aus Kaufpreis, Anteil G+B, AfA-Satz und § 6b-Übertragung.

Das Prognoseblatt hat je Objektzeile einen Block mit 20 Jahreszeilen: Miete, weitere Einnahmen, Erhaltung und weitere Ausgaben wachsen mit ihren Steigerungsraten. Die AfA beträgt AK Gebäude × Satz, höchstens aber den Restbuchwert. Danach sind AfA und Buchwert null. Hat das Objekt Anlagen im Blatt Anlagen, ist die AfA die Summe der AfA je Anlage: Jede Anlage läuft für sich aus, sobald ihr Buchwert verbraucht ist.

Das Blatt **Übersicht** öffnet als erstes. Es zeigt von 2026 bis 2046 den Wert der Immobilien und das Gesamtvermögen am Jahresende, je als Tabelle und Liniendiagramm, und zwar in zwei Linien:

- **Baseline:** alles halten, nichts verkaufen.
- **Plan:** mit den Verkäufen und Neuobjekten aus den Blättern Verkäufe und Neuobjekte.

Der Wert ist der Verkehrswert aus dem Objektblatt, fortgeschrieben mit der Wertsteigerung vom Parameterblatt. Fehlt der Verkehrswert, gilt die Annahme Jahresmiete × Vervielfältiger; die Übersicht zählt die Objekte mit Annahmen. Ein Verkauf gilt zum Jahresende. Miete und AfA laufen im Verkaufsjahr noch, ab dem Folgejahr ist das Objekt inaktiv. Der Plan enthält auch die Neuobjekte. Das Gesamtvermögen ist der Verkehrswert plus die kumulierte Liquidität nach Steuern, abzüglich der Restschuld der Kredite und der latenten Steuer auf stille Reserven und Rücklage. So stehen Halten und Verkaufen vergleichbar nebeneinander.

Im Blatt **Verkäufe** stehen je Verkauf ObjektID (daneben grau die Bezeichnung kurz aus Zelle C2 des Kostenstellenblatts, ohne Kostenstellenblatt der Objektname), Jahr, Preis, Kosten, optional der Anteil G+B laut Kaufvertrag und § 6b ja/nein. Daraus rechnet das Blatt:
- den Gebäudebuchwert am Ende des Verkaufsjahrs aus der Prognose
- die Aufteilung des Nettoerlöses auf Gebäude und G+B
- den Gewinn je Teil

Die Aufteilung folgt dem Kaufvertrag, sonst dem Verkehrswertanteil aus dem Objektblatt. Laut Kaufvertrag kann G+B als Anteil oder als fester Betrag („Preis G+B lt. Kaufvertrag“) angegeben werden. Der Betrag geht vor, das Gebäude erhält den Rest des Verkaufspreises, und die Verkaufskosten werden anteilig aufgeteilt. Der Status meldet unter anderem eine fehlende Aufteilung und eine zu kurze Vorbesitzzeit für § 6b.

Kostenstellen, die zusammen zu einem Gesamtpreis verkauft werden, stehen als **Verkaufspaket** in einer Zeile. Die ObjektID ist die erste Kostenstelle, die übrigen kommen in die Spalte „weitere Kostenstellen“, getrennt durch Semikolon oder Komma (z. B. „KSt 4; KSt 5“).
- **Was für das Paket gilt:** Verkaufspreis, Kosten und Preis bzw. Anteil G+B laut Kaufvertrag gelten für das ganze Paket.
- **Was summiert wird:** Buchwert Gebäude, AK G+B und, für die Preis-Annahme, die Verkehrswerte aller Kostenstellen.
- **Vorbesitzzeit:** Sie zählt ab dem spätesten Kaufjahr im Paket, also streng für § 6b.
- **Prognose:** Alle Kostenstellen des Pakets enden mit dem Verkauf.
- **Rücklage, Neuobjekt und Verkaufsergebnis:** Rücklage und Neuobjekt entstehen einmal für das Paket. Das Verkaufsergebnis erscheint im BWA-Blatt der ersten Kostenstelle.

Das Blatt **Rücklagen** bildet bei § 6b ja und Status OK eine Rücklage aus den positiven Teilgewinnen, getrennt nach Gebäude und G+B. Die Frist beträgt vier Jahre, mit „§ 6b Neubau begonnen = ja“ sechs Jahre. Was bis zum Fristjahr nicht auf Neuobjekte übertragen ist, wird dort aufgelöst, mit 6 % Zuschlag je Jahr. Der Spiegel je Jahr zeigt Gewinne, Einstellung, Übertragung, Auflösung, Zuschlag, Bestand und den steuerwirksamen Betrag, der in Etappe 7 die Steuer ergibt.

Im Blatt **Neuobjekte** stehen je Reinvestition Kaufjahr, Kaufpreis, Anteil G+B (oder AK G+B in Euro: dann darf der Anteil leer bleiben, die Spalte „Anteil G+B verwendet“ zeigt den daraus berechneten Anteil), Nebenkosten, AfA-Satz, AfA-Methode (linear oder degressiv 5 % nach § 7 Abs. 5a EStG; Wechsel zur linearen AfA nur mit Parameter „Wechsel degressiv → linear“ = ja), Mietrendite, Erhaltungsquote und bis zu drei Quell-Rücklagen (Auswahl aus dem Blatt Rücklagen). Gekauft wird zum Jahresende, Miete und AfA laufen ab dem Folgejahr. Die Rücklagen werden in fester Reihenfolge übertragen, über alle Quellen:
1. Gebäudegewinne auf das neue Gebäude
2. G+B-Gewinne auf den neuen G+B (bis auf 0)
3. Rest der G+B-Gewinne auf das Gebäude

**Verkauf ein- und ausschalten:** Mit „Verkauf aktiv = nein“ im Blatt Verkäufe rechnet ein Verkauf nicht mit, die Zeile bleibt aber erhalten. Das Objekt wird dann weiter gehalten (Miete, AfA, Darlehen laufen weiter), es entsteht kein Veräußerungsgewinn und keine § 6b-Rücklage. Ein Neuobjekt, das diese Rücklage als Quelle nennt, überträgt nichts und meldet im Status „Verkauf ausgeschaltet, Rücklage entfällt, keine Übertragung“. Seine AfA-Basis ist wieder die volle AK Gebäude. Ein Neuobjekt aus „reinvestieren = ja“ verschwindet mit dem Verkauf. Der Status des Verkaufs lautet „ausgeschaltet“ und zählt nicht als Fehler, die Zeile erscheint grau durchgestrichen. Ausgeschaltete Zeilen zählen bei „Objekt mehrfach verkauft“ nicht mit, so lassen sich zwei Varianten nebeneinander halten.

**Kauf ein- und ausschalten:** Mit „Kauf aktiv = nein“ im Blatt Neuobjekte rechnet ein Kauf nicht mit, die Zeile bleibt aber erhalten. Er bringt dann keine Miete, keine AfA, keinen Kredit, keine Übertragung der Rücklage und keinen Abfluss in der Liquidität. Eine Rücklage, die er nutzen sollte, wird zum Fristjahr aufgelöst. Der Status lautet „ausgeschaltet“ und zählt nicht als Fehler, die Zeile erscheint grau durchgestrichen. Das Blatt einer gewählten Kostenstelle Neukauf zeigt in den Planjahren 0 und oben den Hinweis. Mit „ja“ oder leer rechnet der Kauf wieder mit und die Kostenstelle füllt sich wieder.

**Miete nach Quadratmetern:** Statt der Mietrendite lassen sich je Neuobjekt Fläche (m²) und Mietpreis je m² (monatlich, zum Kaufzeitpunkt) angeben. Dann gilt Jahresmiete = Fläche × Preis × 12, ab dem Folgejahr mit der Mietsteigerung. Instandhaltung (1250) und Verwaltung (1260, weitere Ausgaben) rechnen in % der Miete des Jahres, statt mit der Erhaltungsquote und ohne Anlauffaktor. Die Sätze stehen auf dem Parameterblatt, Standard 7 % und 2 %. Wählt das Neuobjekt eine Kostenstelle Neukauf, füllen diese Annahmen dort nur die leeren Jahre (Blatt Neukauf-KSt, fortgeschrieben, und das Blatt der Kostenstelle). Eingetragene Werte bleiben, und Jahre nach einem eingetragenen Wert schreiben diesen fort.

**Sonder-AfA § 7b (QNG):** Mit „Sonder-AfA § 7b (QNG) = ja“ zieht ein Neuobjekt zusätzlich zur regulären AfA 5 % der Bemessungsgrundlage ab, im Kaufjahr und den drei Folgejahren (Parameter „Sonder-AfA § 7b“, eingeklappter Abschnitt). Der Betrag ist fest, auch im Kaufjahr voll, und wird vor der regulären AfA gezogen; diese rechnet nur auf den Rest, degressiv vom geminderten Restbuchwert. Nach den vier Jahren endet die Sonder-AfA, linear gilt dann Restwert × AfA-Satz (§ 7a Abs. 9 EStG). Bemessungsgrundlage ist die AfA-Basis (nach § 6b), mit Fläche höchstens 4.000 € je m²; liegen die AK Gebäude über 5.200 € je m², entfällt die Sonder-AfA und der Status meldet es. Die Prognose zeigt sie in „davon Sonder-AfA § 7b“.

**Kauf im Vorjahr der Veräußerung (§ 6b Abs. 1 Satz 1):** Ein Neuobjekt darf ein Jahr vor dem Verkauf seiner Quelle gekauft sein (Parameter „§ 6b Abzug bei Kauf im Vorjahr“ = ja). Der Gewinn wird dann erst im Verkaufsjahr abgezogen, in derselben Reihenfolge ü1, ü2, ü3. Auf das Gebäude geht höchstens der Buchwert am Ende des Kaufjahrs, also AK Gebäude abzüglich der Sonder-AfA § 7b des Kaufjahrs (§ 6b Abs. 5). Im Kaufjahr stehen Gebäude und G+B mit den vollen AK. Im Verkaufsjahr sinken die Buchwerte um den Abzug, und schon die AfA dieses Jahres rechnet von den geminderten AK, ebenso die Sonder-AfA. Im Rücklagenspiegel stehen Einstellung und Übertragung im Verkaufsjahr, es bleibt kein Bestand. Der Verkaufserlös fließt erst nach dem Kauf und mindert den Finanzierungsbedarf nicht; die Prüfung nennt solche Käufe als Hinweis. Ein Kauf zwei oder mehr Jahre vor dem Verkauf überträgt nichts.

Die AfA läuft von der geminderten AfA-Basis. Die Werte je Quelle stehen in eingeklappten Spalten rechts. Neuobjekte erscheinen in der Prognose und in der Plan-Linie der Übersicht, nicht in der Baseline.

**Finanzierung je Neuobjekt:** Die Rücklage wirkt nur steuerlich (Übertragung auf die AK, AfA-Basis); die Finanzierung legt man getrennt fest. Finanzierungsbedarf = Kaufpreis + Nebenkosten. Nur mit „Verkaufserlös zur Finanzierung = ja“ wird zusätzlich der Nettoerlös der Quell-Verkäufe abgezogen (in Zeilenreihenfolge, soweit nicht schon für ein Objekt darüber eingesetzt); leer oder „nein“ lässt den Bedarf bei der ganzen AK. Reinvestitionen aus „reinvestieren = ja“ sind mit „ja“ vorbelegt. Mit „Finanzierung Rest = Kredit“ deckt ein Kredit den Bedarf (Kreditbetrag leer) oder den eingetragenen Betrag, der Rest kommt aus Eigenmitteln (Liquidität). Alternativ gibt die Kreditquote den Kredit in % von Kaufpreis + Nebenkosten vor, z. B. 70 % bei 30 % Eigenkapital. Kreditbetrag und Kreditquote stehen fest: Mit „Verkaufserlös zur Finanzierung = ja“ deckt der Erlös dann nur den Eigenkapitalteil (bei 70 % Quote also 30 % der AK). Was vom Erlös übrig bleibt, steht späteren Käufen derselben Quelle zur Verfügung und bleibt sonst in der Liquidität (verzinst mit der Alternativrendite); versteuert wird nie der Erlös, sondern nur der Veräußerungsgewinn, sofort oder über die Rücklage. Eingaben: Zinssatz, Tilgungsart (Annuität mit anfänglicher Tilgung, linear, endfällig), Tilgung p. a., optional die Laufzeit, nach der die Restschuld auf einmal getilgt wird, und tilgungsfreie Jahre (in den ersten Jahren nach dem Kauf nur Zins). **Zweitkredit:** Ein Teil des Kredits kann zu eigenen Konditionen laufen, z. B. ein günstiger Förderkredit. „Zweitkredit Anteil“ (in % des Kredits) oder „Zweitkredit Betrag“ (fest, geht dem Anteil vor, höchstens der Kredit) legt ihn fest, dazu der Zinssatz (Pflicht). Tilgungsart, Tilgung, Laufzeit und tilgungsfreie Jahre des Zweitkredits übernehmen leer die des Hauptkredits. Der Kredit insgesamt (Finanzierungsbedarf, Eigenmittel, Liquidität) ändert sich dadurch nicht, nur seine Aufteilung: Hauptkredit = Kredit − Zweitkredit. Das Blatt **Darlehen** rechnet je Kredit Zins, Tilgung und Restschuld je Jahr (oben die Hauptkredite, darunter die Zweitkredite); Auszahlung zum Ende des Kaufjahrs, Zins und Tilgung ab dem Folgejahr, nach der vollen Tilgung entfällt beides. Der Zins mindert das Ergebnis (BWA 1310 im Blatt des Neuobjekts und in „Alle Objekte“), die Tilgung nur die Liquidität, die Restschuld das Gesamtvermögen. Szenario C finanziert gleich, in B entfällt mit dem Kauf auch der Kredit.

**Darlehen der Bestandsobjekte:** Das Blatt Objekte hat die Spalten Zinsaufwand Basisjahr (BWA 1310, mit dem Kostenstellenblatt verknüpft), Restschuld Darlehen Ende Basisjahr, Zinssatz und Rate p. a. (Zins + Tilgung).
- Ohne Restschuld wird nur der Zinsaufwand fortgeschrieben, mit „Zinsaufwand ohne Restschuld: Veränderung p. a.“ vom Parameterblatt (Standard −3 %), bis zum Verkaufsjahr. Tilgung und Ablösung fehlen dann; die Prüfung meldet solche Objekte (Hinweis, bei Verkauf Warnung).
- Mit Restschuld: Zins = Restschuld am Vorjahresende × Zinssatz, Tilgung = Rate − Zins. Zinssatz leer: Zinsaufwand / Restschuld; Rate leer: Restschuld × (Zinssatz + Tilgung in % vom Parameterblatt). Im Verkaufsjahr wird die Restschuld aus dem Erlös abgelöst (Tilgung), eine Vorfälligkeitsentschädigung gehört in die Verkaufskosten.
- Der Zins mindert Ergebnis und Steuer (BWA 1310 im Blatt des Objekts und in „Alle Objekte“), Tilgung und Ablösung die Liquidität, die Restschuld das Gesamtvermögen. Die Baseline rechnet die Darlehen weiter, als würde nie verkauft.

Das Blatt **Liquidität** rechnet je Jahr die Steuer und den Geldfluss, je Szenario eine Tabelle:
- Steuer = (laufendes Ergebnis + steuerwirksamer Betrag aus dem Rücklagenspiegel) × Grenzsteuersatz
- Verluste werden vorgetragen und mit späteren Gewinnen verrechnet
- freier Mittelzufluss = Mieten und Einnahmen − Erhaltung und Ausgaben + Zins − Zinsen Darlehen + Verkaufserlöse − Steuer − Kauf der Neuobjekte + Kreditauszahlung − Tilgung (mit Ablösung beim Verkauf)

Die Liquidität liegt in einer Alternativanlage und wird mit der Rendite vom Parameterblatt verzinst (Standard 3 %); der Zins ist steuerpflichtig.

Das Blatt **Auswertung** zeigt je Jahr die Gesamt-GuV, Steuer, kumulierte Steuer, stille Reserven, Rücklagenbestand, Verlustvortrag und Gesamtvermögen vor und nach latenter Steuer. Die latente Steuer ist die Steuer, die anfiele, wenn alle Objekte zum Verkehrswert verkauft und die Rücklage aufgelöst würden.

Alle Szenarien rechnen gleichzeitig, Liquidität und Auswertung haben je Szenario eine Tabelle:
- **A Plan:** § 6b-Kette wie erfasst
- **B:** jeder Veräußerungsgewinn sofort versteuert, Neuobjekte mit Quelle-Rücklage entfallen, das Geld bleibt in der Alternativanlage
- **C:** sofort versteuert, die Neuobjekte werden trotzdem gekauft, mit voller AfA-Basis
- **Baseline:** alles halten

Das Blatt **Vergleich** stellt die Kennzahlen am Ende des Rasters nebeneinander, mit den Differenzen A − B, A − C und A − Baseline, dazu das Endvermögen nach latenter Steuer je Jahr als Diagramm. A − C zeigt die reine Wirkung von § 6b: ohne Alternativrendite ein Gleichstand, mit Rendite der Zins auf die gestundete Steuer.

Das Blatt **Ergebnisbrücke** zeigt, wie die Entscheidungen auf das Ergebnis wirken. Die Spalten sind die Planjahre und eine Summe. Die Zeilen führen vom Ergebnis vor Steuern bei Halten (Baseline) zum Ergebnis vor Steuern im Plan A:
- **Ergebnis bei Halten**, aufgeteilt in Immobilien, Kostenstellen der Art Unternehmen und Zinsertrag der Liquidität.
- **Je Verkauf ein Paket mit Zwischensumme**, darunter zwei Zeilen:
  - **Verkauf:** im Verkaufsjahr der steuerwirksame Gewinn, bei § 6b minus die Einstellung in die Rücklage; im Fristjahr Auflösung und Zuschlag; ab dem Folgejahr minus das laufende Ergebnis, das mit den verkauften Kostenstellen entfällt (Miete − Kosten − AfA − Darlehenszinsen). Ein Verkaufspaket steht in einer Zeile.
  - **Käufe aus dem Verkauf:** alle Neuobjekte, deren erste Quelle mit einem Verkauf (Quelle 1, sonst 2, sonst 3) die Rücklage oder den Erlös dieses Verkaufs nennt. Auch ein Verkauf ohne § 6b zählt, über „RL-“ & ObjektID.
- **Käufe ohne Verkauf:** Neuobjekte ohne Quelle.
- **Summe Pakete und Käufe ohne Verkauf**, davon Gewinn und Rücklage, entfallendes laufendes Ergebnis und Käufe.
- **Käufe einzeln** (unten): je Kauf Miete − Kosten − AfA − Kreditzinsen, mit dem Paket, zu dem er zählt. Die AfA rechnet von den um die Rücklage geminderten AK, mit Sonder-AfA § 7b.
- **Zinseffekt der Liquidität:** Verkaufserlöse, Kaufpreise und Steuer verändern die Liquidität und damit deren Verzinsung.
- **Kontrolle:** Halten + Summe Pakete und Käufe ohne Verkauf + Zinseffekt − Plan, muss 0 sein; sonst meldet das Prüfungsblatt einen Fehler.

Die Steuer steht nachrichtlich darunter (Halten, Plan und die Wirkung nach Steuern), denn mit Verlustvortrag lässt sie sich nicht einzelnen Entscheidungen zuordnen.

**Kostenstellen ohne Immobilie:** Jede Kostenstelle der Kostenstellen-Datei wird ein Objekt, auch eine für gesonderte Unternehmenseinnahmen und -ausgaben. Solche Kostenstellen bekommen im Blatt Objekte die Art „Unternehmen“ (leer = Immobilie). Sie werden wie ein Objekt aus dem Basisjahr fortgeschrieben (Einnahmen mit der Mietsteigerung, Ausgaben mit der Kostensteigerung, AfA aus den Anlagen) und zählen in Liquidität, Steuer und „Alle Objekte“ mit. Sie haben keinen Verkehrswert, und die Ergebnisbrücke weist sie getrennt aus. Ein Verkauf einer solchen Kostenstelle meldet das Prüfungsblatt als Warnung.

Das Blatt **Prüfung** rechnet alle Plausibilitätsprüfungen als Formeln, je mit Art und Anzahl betroffener Zeilen:
- **Fehler:** Status ungleich OK in Objekten, Verkäufen oder Neuobjekten (darunter Fristverstoß), Steuerwelt nicht GmbH
- **Warnung:** Vorbesitzzeit für § 6b zu kurz, Rücklage nur teilweise oder gar nicht übertragen, Drei-Objekt-Grenze überschritten
- **Warnung** außerdem: verkauftes Objekt mit Zinsaufwand, aber ohne Restschuld (Ablösung fehlt)
- **Hinweis:** Frist nach Prognoseende, Liquidität negativ (Finanzierungsbedarf), Kredit am Ende des Rasters nicht getilgt, Zinsaufwand ohne Restschuld grob fortgeschrieben, Verkehrswert fehlt

Das Gesamtergebnis steht auch auf dem Parameterblatt und in der Übersicht.

Mit `--makros` entsteht eine .xlsm mit VBA-Steuerung, die Schaltflächen liegen auf dem Parameterblatt:
- Plausibilität prüfen
- Neu berechnen
- Objekt anlegen, duplizieren, entfernen
- leere Prognoseblöcke aus- und einblenden
- Variante festhalten

„Variante festhalten“ schreibt die Kennzahlen des Vergleichs als feste Werte ins Blatt **Varianten**, so lassen sich etwa verschiedene Verkaufsjahre vergleichen. VBA rechnet nichts, alle Ergebnisse entstehen in den Formeln.

Die Ergebnisse stehen auch im **DATEV-BWA-Format** (Projektplan, Abschnitt 18):
- **Blatt je Kostenstelle:** je Objekt und Neuobjekt ein Blatt mit den BWA-Zeilen 1010–1380. Links stehen die Ist-Werte aus der eingelesenen BWA, rechts die Planjahre als Formeln aus der Prognose. Das Kostenstellenblatt führt: Die Spalte Basisjahr ist Eingabe (gelb). Miete, Erhaltung, weitere Einnahmen und Ausgaben und AfA lt. Buchhaltung im Blatt Objekte verweisen darauf (grün), das Summenblatt „Alle Objekte“ addiert die Kostenstellenblätter. Eine Änderung im Kostenstellenblatt geht so in Objekte, Prognose und „Alle Objekte“ ein. Wird ein solcher Wert im Blatt Objekte überschrieben, rechnet nur die Prognose damit: Die Zelle wird orange, das Prüfungsblatt warnt, und das Makro „Objekte -> Kostenstellen“ schreibt den Wert ins Kostenstellenblatt (weitere Ausgaben: Differenz auf 1260) und stellt die Verknüpfung wieder her. Objekte ohne eingelesene Kostenstelle zeigen im Basisjahr die Werte des Blatts Objekte (hellblau). Die Monatsspalten des Basisjahrs sind gruppiert und eingeklappt; das „+“ über der Jahresspalte öffnet sie.
- **Kostenarten:** Die weiteren Ausgaben werden nach dem Anteil der Kostenart im Basisjahr aufgeteilt.
- **Verkauf und Rücklage:** Gewinn, Einstellung und Auflösung stehen im neutralen Ergebnis.
- **Summenblatt „Alle Objekte“:** Es trägt dazu Zinsertrag und Steuer und stimmt mit Liquidität und Auswertung überein.
- **Blatt „Alle Objekte halten“:** dieselbe Summen-BWA als reine Halten-Ansicht (Baseline): alle Bestandsobjekte über das ganze Raster gehalten, ohne Verkäufe und Neuobjekte, Bestandsdarlehen laufen weiter. Ist-Spalten wie in „Alle Objekte“, Zinsertrag und Steuer aus der Baseline der Liquidität; die Kontrolle im Blatt BWA-Zuordnung prüft auch dieses Blatt.
- **Zahlenformat:** Beträge in ganzen Euro. Die Zellen behalten den genauen Wert, gerundet wird nur die Anzeige.
- **Vergleich mit Halten:** In „Alle Objekte“ steht in D3 der Schalter „Vergleich mit Halten“ (ja/nein). Bei „ja“ zeigen die Spalten rechts neben den Planjahren je Position und Jahr die Differenz Plan − Halten (grün mehr, rot weniger), auch für Herleitung und Finanzierung; bei „nein“ bleiben sie leer. `--uebernehmen` übernimmt den Schalter.
- **Blatt Verkauf und Kauf:** je Verkauf eine Ergebnissicht für die Berichterstattung und eine Detailsicht. Die Ergebnissicht zeigt Erlös, Reinvestition, Kapitalanlage, Übertrag § 6b und den Vergleich Halten gegen Alternative im ersten vollen Jahr. Die Detailsicht zeigt die Einzelauflistung nach G+B und Gebäude und die Planung. Dazu kommt je Neuobjekt die Detailsicht des Kaufs.

**Mandanten-Export:** `python -m prognosemodell.export Prognose.xlsx` schreibt daneben `Prognose_Mandant.xlsx` (anderer Name mit `--ausgabe`), eine schlanke Fassung für den Mandanten, nur mit Werten. Die Mappe muss in Excel gespeichert sein; eine frisch erzeugte, nie geöffnete Mappe rechnet LibreOffice vorher durch. Bis zu sieben Blätter:
- **Überblick:** Empfehlung, Endvermögen je Option, Datenlage, Verkehrswert und Gesamtvermögen je Jahr mit Diagrammen.
- **Jahresrechnung:** je Jahr vom Ergebnis über Steuer und Geldfluss zum Vermögen, für Plan und Halten, dazu die Differenz. Jede Spalte rechnet von oben nach unten.
- **Ergebnisbrücke:** vom Ergebnis bei Halten über die Pakete (Verkauf mit seinen Käufen) zum Ergebnis im Plan, ohne leere Pakete und Vorgänge und ohne Kontrollzeile. Ohne Verkauf und Kauf entfällt das Blatt.
- **BWA gesamt:** Summen-BWA mit Ist-Jahren und Planjahren und der Herleitung der Sonderposten; bei „Vergleich mit Halten = ja“ auch Halten und Differenz.
- **BWA je Objekt:** alle Kostenstellenblätter untereinander, mit Inhaltsverzeichnis.
- **Verkauf und Kauf:** Verkäufe, Neuobjekte, Ergebnis- und Detailsicht je Vorgang. Ohne Verkauf und Kauf entfällt das Blatt.
- **Annahmen:** zentrale Annahmen, Daten je Objekt mit Herkunft (blau Annahme, grün Buchhaltung oder Anlagenverzeichnis), auffällige Prüfungen.

Zeilen ohne Wert, Monatsspalten, Hilfszeilen und ausgeschaltete Vorgänge entfallen; Ist-Vorjahre erscheinen nur, wenn sie eingelesen sind. Alle Blätter sind für den Druck im Querformat eingerichtet.

**Startblatt und Datenlage**

Das Blatt **Start** öffnet als erstes. Es zeigt:
- die Handlungsempfehlung: die Option mit dem höchsten Endvermögen nach latenter Steuer (halten, § 6b-Kette, sofort versteuern und reinvestieren, sofort versteuern und anlegen), mit dem Vorsprung gegenüber Halten;
- den Wert der § 6b-Kette und den tiefsten Liquiditätsstand;
- die Belastbarkeit des Ergebnisses;
- eine Anleitung mit Links, die Farblegende und die zentralen Annahmen.

Fehlende Daten füllt eine **Auffülllogik**: Die Annahme steht als Formel in der leeren Eingabezelle, ihre Sätze stehen zentral auf dem Parameterblatt. Ein eingetippter Wert ersetzt sie.

| Feld | Annahme |
| --- | --- |
| Verkehrswert | Miete × 20 |
| Gebäudeanteil | 75 % |
| AfA-Satz | Blatt Anlagen (AfA p. a. / AK), sonst 2 % |
| Kaufjahr | Blatt Anlagen (frühester Zugang G+B oder Gebäude), sonst vor 15 Jahren |
| AK Gebäude | Blatt Anlagen (AHK der abnutzbaren Anlagen), sonst AfA lt. Buchhaltung / Satz |
| AK G+B | Blatt Anlagen (Buchwert G+B), sonst aus AK Gebäude und Gebäudeanteil |
| Baujahr | Blatt Anlagen (frühestes AHK-Datum der Gebäude, gilt bei Neubau; orange und Warnung, bis das Jahr eingetippt ist), sonst leer: keine Alterung und keine Großmaßnahme |
| Restbuchwert | Blatt Anlagen (Buchwert Ende Basisjahr), sonst aus AK und Kaufjahr; 0, wenn die Buchhaltung keine AfA mehr zeigt |
| AfA je Jahr (Prognose) | Blatt Anlagen (je Anlage bis zu ihrem Buchwert), sonst AfA lt. Buchhaltung (auch 0), läuft bis der Restbuchwert verbraucht ist |
| Erhaltung | 10 % der Miete |
| Verkaufspreis | Verkehrswert fortgeschrieben |

Die **Erhaltung hängt vom Gebäudealter ab**:
- Ab 30 Jahren steigt sie um 1,5 % pro Jahr zusätzlich.
- Neuobjekte tragen in den ersten 10 Jahren nur die Hälfte; Standard 0,5 % des Kaufpreises.
- Ab einem Gebäudealter von 50 Jahren fällt eine Großmaßnahme an, standardmäßig 15 % des Gebäudewerts.

Sie trifft nur das Halten, ein Verkauf davor erspart sie. Baujahr, Jahr und Betrag lassen sich je Objekt überschreiben (Jahr 0 = keine), alle Sätze stehen auf dem Parameterblatt.

„reinvestieren = ja“ im Blatt Verkäufe legt ein Neuobjekt aus den Annahmen an. Pflicht sind nur ObjektID und Miete.

Farben der Eingabezellen:
- **rot:** Pflicht fehlt oder Annahme gelöscht
- **orange:** Annahme bei einem verkauften Objekt; bestimmt Gewinn und Rücklage, daher Warnung
- **blau:** Annahme
- **grün:** aus der Buchhaltung eingelesen oder aus dem Blatt Anlagen; zählt nicht als Annahme
- **gelb:** händisch eingetragen
- **grau:** Formel

Der **Schnellcheck** (`--schnellcheck`) ist dieselbe Rechnung mit schlanker Ansicht: § 6b und Reinvestition stehen als Annahme auf ja, die steuerlichen Stammdaten sind zugeklappt, die Rechenblätter ausgeblendet. Mit ObjektID und Miete je Objekt und einem geplanten Verkauf liefert er eine erste Empfehlung. Die Statusspalte im Objektblatt meldet fehlende Pflichtfelder, doppelte IDs, ein Kaufjahr nach dem Basisjahr und einen Restbuchwert über den AK.

**Anlagenverzeichnis (Blatt Anlagen)**

Zwei Quellen füllen das Blatt, je Kostenstelle gilt die genauere:
- **Anlagenverzeichnis** (`--inventar`, DATEV-Export „Inventarübersicht“): eine Zeile je Anlage mit Konto, AHK-Datum, AHK, Buchwert Wj-Ende, AfA-Art, AfA-% und KOST1. KOST1 verweist auf die Kostenstelle (1 = „KSt 1“). Fehlt KOST1, zählt eine Kostenstellennummer in der Inventarbezeichnung („Grund u. Boden Kostenstelle 5“, „KSt 5“, „Kostenst. 5“); sonst wird die Bezeichnung unscharf mit ObjektID und Name der Kostenstelle verglichen, Wort für Wort („Musterstr.“ passt zu „Musterstraße 1“, „Grund und Boden“ zählt nicht mit). Solche Treffer sind im Blatt Anlagen orange und in der Spalte Zuordnung gekennzeichnet, bitte prüfen. Passen mehrere Objekte gleich gut, bleibt die ObjektID leer. Abgegangene Anlagen entfallen.
- **Aufschlüsselung im Kostenstellenblatt:** unter der BWA die Blöcke „Buchwert, JE“ und „Abschreibungen JW“ (Spalte des Vorjahrs, im Muster F). Je Gruppe entsteht eine Anlage mit Buchwert und Jahres-AfA, aber ohne AHK und Datum. Sie zählt nur für Kostenstellen, die das Anlagenverzeichnis nicht abdeckt; sonst gleicht die Ausgabe beim Einlesen beide Buchwerte ab.

Je Anlage rechnet das Blatt den Buchwert am Ende des Basisjahrs und die AfA je Prognosejahr. Linear gilt AHK × Satz, auf volle Euro aufgerundet wie bei DATEV, bis der Buchwert verbraucht ist; degressiv gilt Satz × Buchwert. Die Art kommt aus Konto (SKR04) und AfA-Art und lässt sich je Zeile ändern:
- **G+B:** Grund und Boden, keine AfA, ergibt AK G+B.
- **Gebäude, BGA, im Bau:** abnutzbar, ergeben AK Gebäude und Restbuchwert. Im Bau ohne AfA, bis Art und Methode nach Fertigstellung umgestellt werden.
- **Finanzanlage, sonstige:** nicht im Modell.

Es zählen nur Zeilen mit Status OK. Anlagen ohne Objekt (keine KOST1 und keine passende Bezeichnung, oder ihre Kostenstelle hat keine Zeile im Blatt Objekte) meldet das Prüfungsblatt als Hinweis. Im Muster betrifft das vor allem Grund und Boden: Ohne KOST1 bleibt AK G+B eine Annahme, bis die ObjektID eingetragen ist. Ein zweiter Hinweis meldet Objekte, deren AfA lt. Anlagen von BWA 1240 abweicht. Die BWA-Blätter je Kostenstelle zeigen unter der Herleitung die Aufschlüsselung „Buchwert, JE“ und „Abschreibungen JW“ je Anlage und Jahr, wie im Kostenstellenblatt der Kanzlei.

Details, Formeln und Testfälle stehen in [docs/Projektplan.md](docs/Projektplan.md).

Eine kompakte Bedienungsanleitung für die Steuerberatung steht in [docs/Bedienungsanleitung_Steuerberatung.docx](docs/Bedienungsanleitung_Steuerberatung.docx) (erzeugt mit `python docs/bedienungsanleitung.py docs/Bedienungsanleitung_Steuerberatung.docx`).

> Alle Steuersätze und Fristen vor dem Echteinsatz mit dem zuständigen Berufsträger prüfen.
