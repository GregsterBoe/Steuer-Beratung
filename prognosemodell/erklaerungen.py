"""Texte für die Tooltips der Mappe (Kommentare an Spaltenköpfen und Parametern).

Eingabefelder erklären sich über Feld.hinweis (modelle.py); hier stehen die
Herleitungen der berechneten Spalten in Worten und die Wirkung der Parameter.
Schlüssel ist der benannte Bereich der Spalte bzw. des Parameters; die Tabellen
in Liquidität und Auswertung gelten für alle Szenarien und sind nach dem
Spaltenschlüssel abgelegt.
"""

# Kernparameter: auf dem Parameterblatt hervorgehoben (★, fett); Abschnitte ohne
# Kernparameter sind eingeklappt
KERNPARAMETER = {
    "par_Basisjahr", "par_Prognosejahre", "par_Steuerwelt", "par_Steuersatz", "par_Mietsteig", "par_NeuMietsteig",
    "par_Erhaltsteig", "par_Wertsteig", "par_Alternativrendite", "par_StatusPruefung",
    "par_AnnVervielfaeltiger", "par_AnnGebaeudeanteil",
}

# Was ein Parameter beeinflusst
PARAMETER_WIRKUNG = {
    "par_Basisjahr": "Letztes Ist-Jahr. Alle Basiswerte im Blatt Objekte gelten für dieses "
                     "Jahr, das Raster beginnt im Folgejahr. Fortschreibungen rechnen ab hier "
                     "(Miete, Verkehrswert, Verkaufspreis-Annahme).",
    "par_Prognosejahre": "Länge des Rasters. Die Blätter sind beim Erzeugen für diese Zahl "
                         "Jahre angelegt; ein anderer Wert braucht eine neu erzeugte Mappe. "
                         "Das Endvermögen im Vergleich gilt für das letzte Jahr.",
    "par_Startjahr": "Berechnet. Erstes Jahr in Prognose, Liquidität, Auswertung, BWA-Blättern.",
    "par_Endjahr": "Berechnet. Verkaufs- und Kaufjahre müssen bis hier liegen; danach "
                   "endende § 6b-Fristen meldet die Prüfung.",
    "par_Steuerwelt": "GmbH: Steuersatz GmbH, Verluste nur vorgetragen. GmbH & Co. KG: Satz "
                      "aus ESt, SolZ und GewSt mit Anrechnung § 35 (Abschnitt GmbH & Co. KG), "
                      "Verluste wahlweise sofort mit anderen Einkünften verrechnet. Andere "
                      "Werte ergeben einen Fehler in der Prüfung.",
    "par_Steuersatz": "Nur GmbH: wird zum angewandten Steuersatz. Steuer je Jahr = "
                      "Bemessungsgrundlage × Satz (Liquidität) und latente Steuer = stille "
                      "Reserven + Rücklage − Verlustvortrag, × Satz (Auswertung). Wirkt damit "
                      "auf Liquidität, Zinsertrag und das Endvermögen aller Szenarien.",
    "par_SteuersatzEff": "Berechnet: der Satz, mit dem Liquidität, Auswertung und BWA rechnen. "
                         "GmbH = Steuersatz GmbH. KG = GewSt je € + (ESt − Anrechnung) × "
                         "(1 + SolZ); Anrechnung = MIN(Faktor × Messzahl, GewSt, ESt).",
    "par_KGESt": "Nur GmbH & Co. KG: Grenzsteuersatz des Gesellschafters auf den Gewinnanteil "
                 "(100 %). Bestimmt den angewandten Satz und die Erstattung bei Verlusten.",
    "par_KGSoli": "Nur GmbH & Co. KG: SolZ auf die ESt nach Anrechnung der GewSt.",
    "par_KGHebesatz": "Nur GmbH & Co. KG: GewSt je € Gewinn = Messzahl × Hebesatz. Bis zum "
                      "Hebesatz Anrechnungsfaktor × 100 % (400 %) neutralisiert § 35 die GewSt.",
    "par_KGMesszahl": "Nur GmbH & Co. KG: Steuermesszahl, gesetzlich 3,5 %.",
    "par_KGAnrechnung": "Nur GmbH & Co. KG: Anrechnung auf die ESt = Faktor × Messbetrag, "
                        "höchstens die gezahlte GewSt.",
    "par_KGKuerzung": "Nur GmbH & Co. KG: ja = keine GewSt und keine Anrechnung; der Satz "
                      "ist dann ESt × (1 + SolZ). Vereinfacht: gilt auch für die "
                      "Veräußerungsgewinne.",
    "par_KGVerlustausgleich": "Nur GmbH & Co. KG: ja = Verluste mindern die ESt im selben Jahr "
                              "(Steuer negativ), der Verlustvortrag gilt nur für die GewSt; "
                              "nein = Verluste werden wie bei der GmbH vorgetragen.",
    "par_KGGewSt": "Berechnet: Messzahl × Hebesatz, mit erweiterter Kürzung 0.",
    "par_VerlustSatz": "Berechnet: Steuer je € Verlust bzw. je € Gewinn, der nur "
                       "gewerbesteuerlich vom Vortrag gedeckt ist. GmbH und KG ohne "
                       "Verlustausgleich: 0.",
    "par_Mietsteig": "Miete und weitere Einnahmen je Jahr = Basiswert × (1 + Satz)^Jahre. "
                     "Gilt für den Bestand; ein Bestandsobjekt kann im Blatt Objekte einen "
                     "eigenen Satz haben. Neuobjekte: Mietsteigerung Neuobjekte.",
    "par_NeuMietsteig": "Mietsteigerung der Neuobjekte, getrennt vom Bestand: Miete und "
                        "weitere Einnahmen ab dem Kauf × (1 + Satz)^Jahre, auch Jahre ohne "
                        "Planwert der Neukauf-Kostenstelle. Ein Neuobjekt kann im Blatt "
                        "Neuobjekte einen eigenen Satz und Jahre ohne Mietsteigerung haben.",
    "par_Erhaltsteig": "Erhaltung je Jahr wächst mit diesem Satz; auch Großmaßnahmen "
                       "(heutige Preise bis zum Jahr der Maßnahme). Ein Bestandsobjekt kann im "
                       "Blatt Objekte einen eigenen Satz haben.",
    "par_Kostensteig": "Weitere Ausgaben je Jahr = Basiswert × (1 + Satz)^Jahre.",
    "par_Wertsteig": "Verkehrswert je Jahr = Verkehrswert aktuell × (1 + Satz)^Jahre. Bestimmt "
                     "die Verkaufspreis-Annahme, stille Reserven, latente Steuer und das "
                     "Endvermögen. Neuobjekte: ab dem Kaufpreis im Kaufjahr.",
    "par_GrESt": "Zur Information; die Nebenkosten der Neuobjekte kommen aus der Eingabe bzw. "
                 "der Annahme „Kaufnebenkosten in % des Kaufpreises“.",
    "par_AfADegressiv": "Neuobjekte mit AfA-Methode degressiv: AfA = Satz × Restbuchwert; "
                        "mit „Wechsel degressiv → linear“ = ja linear, sobald diese höher ist.",
    "par_AfAWechsel": "Neuobjekte mit AfA-Methode degressiv. nein: jedes Jahr Satz × "
                      "Restbuchwert. ja: Wechsel zur linearen AfA (Restbuchwert / "
                      "Restnutzungsdauer, Nutzungsdauer = 1 / AfA-Satz), sobald diese höher ist; "
                      "bei AfA-Satz ≥ degressivem Satz ist das von Anfang an so.",
    "par_QNGSatz": "Neuobjekte mit „Sonder-AfA § 7b (QNG)“ = ja: Sonder-AfA je Jahr = "
                   "Bemessungsgrundlage § 7b × Satz, fest und zusätzlich zur regulären AfA. "
                   "Wird vor der regulären AfA gezogen und mindert deren Restbuchwert.",
    "par_QNGJahre": "Begünstigungszeitraum: Kaufjahr und Folgejahre. Danach nur noch die "
                    "reguläre AfA, linear vom Restwert × AfA-Satz (§ 7a Abs. 9 EStG).",
    "par_QNGHoechstQm": "Nur Neuobjekte mit Fläche: Bemessungsgrundlage § 7b höchstens Fläche × "
                        "Betrag.",
    "par_QNGObergrenzeQm": "Nur Neuobjekte mit Fläche: AK Gebäude / Fläche darüber → keine "
                           "Sonder-AfA, Status meldet es.",
    "par_Alternativrendite": "Zinsertrag je Jahr = Liquidität am Vorjahresende × Satz, voll "
                             "steuerpflichtig; negative Liquidität kostet denselben Satz. "
                             "Entscheidet den Vergleich A gegen B (Immobilie oder Geldanlage).",
    "par_StatusPruefung": "Berechnet: Ergebnis aller Plausibilitätsprüfungen. Ungleich OK: "
                          "Blatt Prüfung zeigt, wo nachzusehen ist.",
    "par_AnlStand": "Jahr der Buchwerte im Blatt Anlagen. Bis zum Basisjahr schreibt das Blatt "
                    "die AfA je Anlage fort.",
    "par_6bVorbesitz": "Verkauf mit § 6b ja: Vorbesitzzeit (Verkaufsjahr − Kaufjahr) darunter "
                       "→ keine Rücklage, Gewinn sofort steuerpflichtig (Status im Blatt "
                       "Verkäufe).",
    "par_6bFrist": "Fristjahr der Rücklage = Verkaufsjahr + Frist. Kauf danach überträgt "
                   "nichts; der Rest wird im Fristjahr aufgelöst.",
    "par_6bFristNeubau": "Ersetzt die Frist, wenn im Blatt Verkäufe „§ 6b Neubau begonnen = ja“.",
    "par_6bVorjahr": "ja: Ein Neuobjekt darf ein Jahr vor dem Verkauf seiner Quelle gekauft "
                     "sein; der Gewinn wird im Verkaufsjahr abgezogen, beim Gebäude höchstens "
                     "vom Buchwert Ende Kaufjahr, und die AfA rechnet schon in diesem Jahr von "
                     "den geminderten AK. nein: nur Kauf ab dem Verkaufsjahr.",
    "par_6bZuschlag": "Gewinnzuschlag = aufgelöster Betrag × Satz × Jahre zwischen Bildung und "
                      "Auflösung; steuerwirksam im Fristjahr (Szenario A).",
    "par_DOGrenze": "Warnung in der Prüfung, wenn mehr Verkäufe im Zeitraum liegen.",
    "par_DOJahre": "Zeitfenster der Drei-Objekt-Grenze.",
    "par_AnnVervielfaeltiger": "Nur ohne eingetragenen Verkehrswert: Verkehrswert = (Miete + "
                               "weitere Einnahmen) × Faktor. Wirkt auf Verkaufspreis-Annahme, "
                               "Veräußerungsgewinn, stille Reserven und Endvermögen.",
    "par_AnnGebaeudeanteil": "Nur ohne eigene Werte: teilt Verkehrswert, Verkaufserlös und AK "
                             "auf Gebäude und G+B. Bestimmt, welcher Teil des Gewinns auf "
                             "Gebäude bzw. G+B fällt (§ 6b-Übertragung, AfA).",
    "par_AnnAfASatz": "Nur ohne eigenen AfA-Satz: AK Gebäude = AfA lt. Buchhaltung / Satz bzw. "
                      "AfA je Jahr = AK × Satz.",
    "par_AnnHaltedauer": "Nur ohne Kaufjahr: Kaufjahr = Basisjahr − Jahre. Wirkt auf die "
                         "§ 6b-Vorbesitzzeit und die abgezinsten AK.",
    "par_AnnErhQuote": "Nur ohne Erhaltung aus der Buchhaltung: Erhaltung = Miete × Quote.",
    "par_AnnReinvestJahre": "Verkäufe mit „reinvestieren = ja“: Kaufjahr des Neuobjekts = "
                            "Verkaufsjahr + Jahre.",
    "par_AnnReinvestQuote": "Verkäufe mit „reinvestieren = ja“: Kaufpreis + Nebenkosten = "
                            "Nettoerlös × Quote.",
    "par_AnnNeuNebenkosten": "Neuobjekte aus „reinvestieren = ja“: Nebenkosten = Kaufpreis × "
                             "Satz.",
    "par_AnnNeuAnteilGuB": "Neuobjekte aus „reinvestieren = ja“: Anteil G+B.",
    "par_AnnNeuAfASatz": "Neuobjekte aus „reinvestieren = ja“: AfA-Satz.",
    "par_AnnNeuAfAMethode": "Neuobjekte aus „reinvestieren = ja“: AfA-Methode.",
    "par_AnnNeuMietrendite": "Neuobjekte aus „reinvestieren = ja“: Miete = Kaufpreis × Rendite.",
    "par_AnnNeuErhQuote": "Neuobjekte aus „reinvestieren = ja“: Erhaltung = Kaufpreis × Quote.",
    "par_QmErhQuote": "Neuobjekte mit Fläche und Mietpreis je m²: Erhaltung (BWA 1250) = Miete "
                      "des Jahres × Quote, statt Erhaltungsquote auf den Kaufpreis. Füllt leere "
                      "Jahre der gewählten Kostenstelle Neukauf; eingetragene Werte gehen vor.",
    "par_QmVerwQuote": "Neuobjekte mit Fläche und Mietpreis je m²: weitere Ausgaben (BWA 1260, "
                       "Verwaltung) = Miete des Jahres × Quote. Füllt leere Jahre der gewählten "
                       "Kostenstelle Neukauf; eingetragene Werte gehen vor.",
    "par_ErhAlterungAb": "Erhaltung aller Objekte mit Baujahr: ab diesem Gebäudealter kommt je "
                         "Jahr die zusätzliche Steigerung dazu.",
    "par_ErhAlterung": "Zusätzliche Steigerung der Erhaltung je Jahr über dem Alter.",
    "par_NeuErhAnlaufJahre": "Neuobjekte: so viele Jahre nach dem Kauf gilt der Anlauffaktor.",
    "par_NeuErhAnlaufFaktor": "Neuobjekte: Erhaltung in den Anlaufjahren = Erhaltung × Faktor.",
    "par_SanAlter": "Nur ohne Eingabe im Blatt Objekte: Jahr der Großmaßnahme = Baujahr + Alter; "
                    "ohne Baujahr keine.",
    "par_SanVorlauf": "Großmaßnahme frühestens im ersten Prognosejahr + Vorlauf.",
    "par_SanQuote": "Nur ohne Eingabe: Großmaßnahme = Verkehrswert × Gebäudeanteil × Quote, "
                    "als Erhaltung sofort abziehbar. 0 % schaltet die Annahme ab.",
    "par_ZinsVeraenderung": "Nur Objekte mit Zinsaufwand, aber ohne Restschuld: Zinsen je "
                            "Jahr = Zinsaufwand Basisjahr × (1 + Satz)^Jahre, bis zum "
                            "Verkaufsjahr. Mindern Ergebnis, Steuer und Liquidität; Tilgung "
                            "und Ablösung fehlen dann.",
    "par_AnnDarlTilgung": "Nur Objekte mit Restschuld, aber ohne Rate: Rate = Restschuld × "
                          "(Zinssatz + Satz). Bestimmt Tilgung, Restschuld und damit die "
                          "Ablösung beim Verkauf.",
}

# Berechnete Spalten, nach benanntem Bereich
SPALTEN = {
    # Objekte: Hilfsspalten nach den Eingaben
    "obj_Status": "OK oder der erste Fehler: Pflichtfeld fehlt, Annahme gelöscht, ObjektID "
                  "doppelt, Kaufjahr nach Basisjahr, Restbuchwert > AK Gebäude. Nur Zeilen mit "
                  "OK zählen in der Prüfung als in Ordnung.",
    "obj_Annahmen": "Anzahl Felder dieser Zeile, in denen noch die Annahmeformel steht (blau).",
    "obj_Kritisch": "Anzahl Annahmen, die bei einem Verkauf den Gewinn bestimmen (orange, "
                    "sobald das Objekt im Blatt Verkäufe steht).",
    "obj_AnlAbn": "Anzahl Anlagen Gebäude, BGA oder im Bau mit dieser ObjektID und Status OK "
                  "im Blatt Anlagen.",
    "obj_AnlAK": "Davon mit AHK; nur dann kommen AK Gebäude und AfA-Satz aus dem Blatt Anlagen.",
    "obj_AnlGuB": "Anzahl Anlagen G+B dieses Objekts; liefert AK G+B.",
    "obj_AnlKauf": "Anzahl Anlagen mit Zugangsjahr; das früheste ist das Kaufjahr.",
    "obj_AnlBau": "Frühestes AHK-Datum eines Gebäudes: Baujahr-Vorschlag, stimmt nur bei Neubau.",
    "obj_AnlBauOffen": "1 = Baujahr stammt noch ungeprüft aus dem Blatt Anlagen (orange).",
    "obj_AnlAfA": "Summe AfA Basisjahr der abnutzbaren Anlagen dieses Objekts.",
    "obj_AnlDiff": "AfA lt. Anlagen − AfA lt. Buchhaltung; über 1 € meldet die Prüfung.",
    # Verkäufe
    "vk_Umfang": "ObjektID und weitere Kostenstellen der Zeile als ;KSt 1;KSt 4; – alle "
                 "Kostenstellen, die mit diesem Verkauf enden (Prognose, Prüfungen).",
    "vk_Vorbesitz": "Verkaufsjahr − Kaufjahr (Blatt Objekte); mit weiteren Kostenstellen ab "
                    "dem spätesten Kaufjahr. Unter der Mindest-Vorbesitzzeit "
                    "(Parameter) ist § 6b nicht möglich.",
    "vk_BuchwertGeb": "Buchwert Gebäude am Ende des Verkaufsjahrs aus dem Blatt Prognose (nach "
                      "der AfA des Verkaufsjahrs); mit weiteren Kostenstellen deren Summe.",
    "vk_AKGuB": "AK G+B aus dem Blatt Objekte = Buchwert G+B (keine AfA); mit weiteren "
                "Kostenstellen deren Summe.",
    "vk_Nettoerloes": "Verkaufspreis − Verkaufskosten.",
    "vk_AnteilGuB": "Preis G+B lt. Kaufvertrag / Verkaufspreis, sonst Anteil G+B lt. "
                    "Kaufvertrag, sonst 1 − Verkehrswertanteil Gebäude aus dem Blatt Objekte. "
                    "Mit weiteren Kostenstellen ist der Verkehrswertanteil Gebäude nach "
                    "Verkehrswert gewichtet.",
    "vk_ErloesGeb": "Nettoerlös − Erlösanteil G+B.",
    "vk_ErloesGuB": "Nettoerlös × Anteil G+B verwendet.",
    "vk_GewinnGeb": "Erlösanteil Gebäude − Buchwert Gebäude Ende Verkaufsjahr. Positiv: "
                    "Rücklage Gebäude bei § 6b; negativ: Verlust wirkt sofort.",
    "vk_GewinnGuB": "Erlösanteil G+B − AK G+B. Positiv: Rücklage G+B bei § 6b.",
    "vk_Gewinn": "Gewinn Gebäude + Gewinn G+B. Ohne § 6b sofort steuerpflichtig im "
                 "Verkaufsjahr; mit § 6b gehen die positiven Teile in die Rücklage.",
    "vk_Miete": "Miete (BWA 1020) im Verkaufsjahr laut Prognose, das letzte Jahr mit Miete; "
                "mit weiteren Kostenstellen die Summe des Pakets. Ohne weitere Einnahmen.",
    "vk_Rendite": "Bruttomietrendite = Jahresmiete Verkaufsjahr / Verkaufspreis: die Miete, die "
                  "je Euro Verkaufspreis wegfällt. Vergleichswert zur Mietrendite der "
                  "Neuobjekte; liegt sie dort niedriger, kauft der Erlös weniger Miete.",
    "vk_Faktor": "Kaufpreisfaktor = Verkaufspreis / Jahresmiete, der Kehrwert der "
                 "Mietrendite (20-fach = 5 %).",
    "vk_MieteOE": "Miete bei Halten, Durchschnitt der 10 Jahre nach dem Verkaufsjahr (soweit "
                  "im Raster): die Miete, die mit dem Verkauf entfällt.",
    "vk_ErhOE": "Erhaltung bei Halten, Durchschnitt der 10 Jahre nach dem Verkaufsjahr, mit "
                "Alterung und Großsanierung: die Erhaltung, die mit dem Verkauf entfällt.",
    "vk_ErhAnteil": "Erhaltung Ø / Miete Ø der 10 Jahre bei Halten. Vergleichswert zur "
                   "Erhaltungsquote der Neuobjekte; liegt sie hier höher, spart der Tausch "
                   "Erhaltung.",
    "vk_Status": "OK oder der erste Fehler (ObjektID unbekannt, doppelt verkauft, Jahr "
                 "fehlt/außerhalb, Preis fehlt, Aufteilung fehlt). „§ 6b unzulässig“: die Zeile "
                 "rechnet, der Gewinn wird aber sofort versteuert. Andere Fehler: die Zeile "
                 "rechnet nicht mit. „ausgeschaltet“ (Verkauf aktiv = nein): rechnet nicht mit, "
                 "kein Fehler; zählt bei „mehrfach verkauft“ nicht.",
    "vk_PreisAnnahme": "1 = Verkaufspreis ist noch die Annahme (Verkehrswert fortgeschrieben).",
    # Rücklagen je Verkauf
    "rl_ID": "„RL-“ + ObjektID; gefüllt nur bei § 6b ja und Status OK. Diese ID wählt das "
             "Blatt Neuobjekte als Quelle.",
    "rl_ObjektID": "Verkauftes Objekt.",
    "rl_Jahr": "Verkaufsjahr; die Rücklage entsteht zum Jahresende.",
    "rl_Geb": "Positiver Gewinn Gebäude (Blatt Verkäufe), sonst 0.",
    "rl_GuB": "Positiver Gewinn G+B, sonst 0.",
    "rl_Betrag": "Rücklage Gebäude + Rücklage G+B.",
    "rl_Fristjahr": "Bildungsjahr + Reinvestitionsfrist (Parameter; bei Neubau begonnen die "
                    "längere Frist). Spätester Kauf.",
    "rl_UebGeb": "Summe ü1 aller Neuobjekte, die diese Rücklage nennen (Gebäudegewinn auf "
                 "Gebäude).",
    "rl_UebGuB": "Summe ü2 + ü3 aller Neuobjekte, die diese Rücklage nennen (G+B-Gewinn auf "
                 "G+B bzw. Gebäude).",
    "rl_Aufloesung": "Rücklage − übertragen; wird im Fristjahr gewinnerhöhend aufgelöst.",
    "rl_Zuschlag": "Auflösung × Gewinnzuschlag × (Fristjahr − Bildungsjahr); steuerwirksam im "
                   "Fristjahr.",
    "rl_Hinweis": "Frist endet nach dem Prognoseende: Auflösung und Zuschlag fehlen im Raster.",
    # Rücklagenspiegel je Jahr
    "rls_Gewinne": "Veräußerungsgewinne aller gültigen Verkäufe dieses Jahres (Blatt Verkäufe).",
    "rls_Bildung": "Summe der Rücklagen mit diesem Bildungsjahr.",
    "rls_Uebertragung": "Summe „übertragen gesamt“ der Neuobjekte mit diesem Kaufjahr.",
    "rls_Aufloesung": "Summe der Auflösungen mit diesem Fristjahr.",
    "rls_Zuschlag": "Summe der Gewinnzuschläge mit diesem Fristjahr.",
    "rls_Steuerwirksam": "Gewinne − Einstellung + Auflösung + Zuschlag: geht in Szenario A in "
                         "die Steuer.",
    "rls_Bestand": "Vorjahr + Einstellung − Übertragung − Auflösung.",
    # Neuobjekte
    "ne_Gueltig": "1 = rechnet im Modell: Pflichtfelder und Anteil G+B oder AK G+B da, NeuID "
                  "eindeutig und keine ObjektID, "
                  "Kaufjahr im Raster.",
    "ne_AKGuBNeu": "Eingabe AK G+B; leer: (Kaufpreis + Nebenkosten) × Anteil G+B.",
    "ne_AKGebNeu": "Eingabe AK Gebäude; leer: Kaufpreis + Nebenkosten − AK G+B neu.",
    "ne_AnteilGuBVerw": "AK G+B neu / (Kaufpreis + Nebenkosten): der eingetragene Anteil G+B "
                        "oder, bei AK G+B in Euro, der daraus berechnete Anteil.",
    "ne_AKAbweichung": "AK G+B neu + AK Gebäude neu − (Kaufpreis + Nebenkosten); ungleich 0 "
                       "nur, wenn beide Beträge von Hand eingetragen sind (Prüfung warnt).",
    "ne_RLGeb": "Summe der Quellen: Rücklage Gebäude, soweit nicht schon von Zeilen darüber "
                "verwendet.",
    "ne_RLGuB": "Summe der Quellen: Rücklage G+B, soweit nicht schon verwendet.",
    "ne_Ue1": "Gebäudegewinn auf das Gebäude: höchstens AK Gebäude neu.",
    "ne_Ue2": "G+B-Gewinn auf G+B: höchstens AK G+B neu.",
    "ne_Ue3": "Rest des G+B-Gewinns auf das Gebäude: höchstens AK Gebäude neu − ü1.",
    "ne_UeGesamt": "ü1 + ü2 + ü3: mindert die AK und geht als Übertragung in den "
                   "Rücklagenspiegel.",
    "ne_UeFolge": "Teil der Übertragung aus Quellen, die erst im Jahr nach dem Kauf verkauft "
                  "werden (§ 6b Abs. 1 Satz 1): wird im Verkaufsjahr abgezogen, bis dahin "
                  "laufen Buchwert und Sonder-AfA ohne ihn.",
    "ne_UeFolgeGeb": "davon ü1 + ü3 auf das Gebäude; höchstens Buchwert Ende Kaufjahr "
                     "(AK Gebäude − Sonder-AfA § 7b des Kaufjahrs, § 6b Abs. 5).",
    "ne_QNGBasisKauf": "Bemessungsgrundlage § 7b im Kaufjahr: ohne den erst im Verkaufsjahr "
                       "abgezogenen Gewinn; ab dann gilt die geminderte (§ 7a Abs. 1 Satz 3).",
    "ne_AfABasis": "AK Gebäude neu − ü1 − ü3: Basis für AfA und Buchwert Gebäude.",
    "ne_AKGuB": "AK G+B neu − ü2: Buchwert G+B.",
    "ne_QNGBasis": "Nur mit Sonder-AfA § 7b (QNG) = ja: AfA-Basis Gebäude, mit Fläche höchstens "
                   "Fläche × Höchstbetrag je m²; 0 über der Baukostenobergrenze.",
    "ne_QNGBasisOhne6b": "Wie Bemessungsgrundlage § 7b, aber von der vollen AK Gebäude "
                         "(Szenarien B und C).",
    "ne_MitQuelle": "1 = mindestens eine Quelle angegeben (Quellen ausgeschalteter Verkäufe "
                    "zählen nicht); das Neuobjekt entfällt in Szenario B.",
    "ne_AKGesamt": "Kaufpreis + Kaufnebenkosten.",
    "ne_Erloes": "Nettoerlös der Quell-Verkäufe, der in diesen Kauf fließt (höchstens Kaufpreis "
                 "+ Nebenkosten − fester Kredit aus Kreditbetrag oder -quote, soweit nicht "
                 "schon von Zeilen darüber eingesetzt). Der Rest bleibt in der Liquidität. 0 ohne "
                 "Verkaufserlös zur Finanzierung = ja: die Rücklage wirkt dann nur steuerlich.",
    "ne_Bedarf": "Kaufpreis + Nebenkosten − Einsatz Verkaufserlös: was anders finanziert werden "
                 "muss.",
    "ne_Kredit": "Bei Finanzierung Rest = Kredit: Kreditbetrag; sonst (Kaufpreis + "
                 "Nebenkosten) × Kreditquote, fest vor dem Verkaufserlös; ohne beides der "
                 "ganze Finanzierungsbedarf. Auszahlung zum Ende des Kaufjahrs, Plan im Blatt "
                 "Darlehen.",
    "ne_Kredit2": "Teil des Kredits zu den Konditionen des Zweitkredits: Betrag Zweitkredit "
                  "(höchstens der Kredit), sonst Kredit × Anteil Zweitkredit; 0 ohne. Der "
                  "Hauptkredit ist der Rest. Beide mit eigenem Plan im Blatt Darlehen.",
    "ne_Eigen": "Finanzierungsbedarf − Kredit: kommt aus der Liquidität.",
    "ne_MieteQmJahr": "Fläche × Mietpreis je m² × 12, leer ohne beide. Ersetzt Kaufpreis × "
                      "Mietrendite; Erhaltung und Verwaltung dann in % der Miete "
                      "(Parameterblatt).",
    "ne_MietsteigV": "Mietsteigerung des Neuobjekts, leer die Mietsteigerung Neuobjekte "
                     "(Parameter). Gilt für Miete und weitere Einnahmen.",
    "ne_MieteKauf": "Jahresmiete zum Kaufzeitpunkt: Fläche × Mietpreis je m² × 12, sonst "
                    "Kaufpreis × Mietrendite; mit Kostenstelle Neukauf deren Miete im ersten "
                    "Jahr nach dem Kauf. Ohne weitere Einnahmen.",
    "ne_Rendite": "Bruttomietrendite = Jahresmiete bei Kauf / Kaufpreis (ohne Nebenkosten): "
                  "die Miete je Euro Kaufpreis. Vergleichswert zur Mietrendite der Verkäufe.",
    "ne_Faktor": "Kaufpreisfaktor = Kaufpreis / Jahresmiete bei Kauf, der Kehrwert der "
                 "Mietrendite (20-fach = 5 %).",
    "ne_RenditeQuelle": "Bruttomietrendite der Verkäufe, deren Rücklage die Quellen nennen: "
                        "Summe Jahresmiete Verkaufsjahr / Summe Verkaufspreis. Höher als die "
                        "Mietrendite des Neuobjekts: der Tausch bringt weniger Miete je Euro.",
    "ne_MieteOE": "Miete laut Prognose, Durchschnitt der 10 Jahre nach dem Kaufjahr (soweit im "
                  "Raster); leer, solange das Neuobjekt nicht im Modell ist.",
    "ne_ErhOE": "Erhaltung laut Prognose, Durchschnitt der 10 Jahre nach dem Kaufjahr, mit den "
                "geminderten Anlaufjahren.",
    "ne_ErhAnteil": "Erhaltung Ø / Miete Ø der 10 Jahre nach dem Kauf. Vergleichswert zur "
                   "Erhaltungsquote der Verkäufe.",
    "ne_ErhAnteilQuelle": "Erhaltung in % der Miete der Verkäufe, deren Rücklage die Quellen "
                         "nennen: Summe Erhaltung Ø / Summe Miete Ø bei Halten.",
    "ne_Status": "OK oder Fehler. Pflichtfeld fehlt, NeuID doppelt bzw. wie Bestandsobjekt, "
                 "Kaufjahr außerhalb Raster: das Neuobjekt rechnet nicht. Meldungen zu Quellen "
                 "(unbekannt, Verkauf ausgeschaltet, Kauf vor Bildung, nach Fristjahr) und Kredit: "
                 "es rechnet, aber ohne diese Übertragung bzw. mit dem gemeldeten Mangel.",
    # Prognose
    "prg_Aktiv": "1 = Objekt erzielt in diesem Jahr Miete und AfA: Bestand bis einschließlich "
                 "Verkaufsjahr, Neuobjekt ab dem Jahr nach dem Kauf.",
    "prg_MieteHalten": "Bestand ohne Verkauf: Miete Basisjahr × (1 + Mietsteigerung des "
                       "Objekts)^Jahre, ab „Miete eingefroren ab“ auf dem Stand des Vorjahrs; "
                       "Neuobjekt 0. Grundlage der Baseline.",
    "prg_EinnahmenHalten": "Bestand ohne Verkauf: weitere Einnahmen Basisjahr × (1 + "
                           "Mietsteigerung des Objekts)^Jahre; Neuobjekt 0.",
    "prg_Miete": "Bestand: Miete bei Halten, bis zum Verkaufsjahr. Neuobjekt: Planwert der "
                 "Neukauf-Kostenstelle, sonst Fläche × Mietpreis je m² × 12 bzw. Kaufpreis × "
                 "Mietrendite, ab dem Kauf mit der Mietsteigerung des Neuobjekts; mit "
                 "„Mietsteigerung erst nach Jahren“ so lange auf dem Anfangsstand.",
    "prg_Erhaltung": "Erhaltung Basisjahr, fortgeschrieben mit Erhaltungssteigerung, Alterung "
                     "über dem Gebäudealter (nur mit Baujahr), plus Großmaßnahme im Jahr der Maßnahme.",
    "prg_AfA": "AfA laut Blatt Anlagen bzw. AfA-Plan, sonst AfA je Jahr bis der Restbuchwert "
               "verbraucht ist. Neuobjekt: Sonder-AfA § 7b + AfA-Basis × Satz (bzw. "
               "degressiv vom Restbuchwert nach der Sonder-AfA).",
    "prg_AfAQNG": "Neuobjekt mit QNG: Bemessungsgrundlage § 7b × Satz im Kaufjahr und den "
                  "Folgejahren (Parameter), vor der regulären AfA; in der AfA enthalten.",
    "prg_AfAQNGOhne6b": "Sonder-AfA § 7b von der Bemessungsgrundlage ohne § 6b (Szenarien B "
                        "und C); in AfA ohne § 6b enthalten.",
    "prg_Verkehrswert": "Verkehrswert aktuell × (1 + Wertsteigerung)^Jahre, unabhängig vom "
                        "Verkauf. Neuobjekt: Kaufpreis × (1 + Wertsteigerung)^Jahre seit Kauf.",
    "prg_Buchwert": "Buchwert Vorjahr − AfA, nicht unter 0.",
    "prg_Ergebnis": "Miete + weitere Einnahmen − Erhaltung − weitere Ausgaben − AfA (mit "
                    "Sonder-AfA § 7b, beim Neuobjekt schon im Kaufjahr).",
    "prg_Bestand": "1 = am Jahresende im Bestand (Verkaufsjahr schon 0).",
    "prg_KreditZins": "Neuobjekt: Zinsen des Kredits (Blatt Darlehen). Bestand mit Restschuld: "
                      "Restschuld Vorjahresende × Zinssatz; ohne Restschuld: Zinsaufwand "
                      "Basisjahr × (1 + Veränderung)^Jahre bis zum Verkauf.",
    "prg_Tilgung": "Neuobjekt: Blatt Darlehen. Bestand: Rate − Zins, höchstens die Restschuld; "
                   "im Verkaufsjahr die ganze Restschuld (Ablösung aus dem Erlös).",
    "prg_Abloesung": "Teil der Tilgung über die planmäßige hinaus: im Verkaufsjahr eines "
                     "Bestandsobjekts die abgelöste Restschuld, sonst 0.",
    "prg_Restschuld": "Restschuld Vorjahresende − Tilgung.",
    "prg_ZinsHalten": "Zinsen, als würde das Bestandsobjekt nie verkauft (Baseline).",
    "prg_TilgungHalten": "Tilgung, als würde das Bestandsobjekt nie verkauft (Baseline).",
    "prg_RestschuldHalten": "Restschuld, als würde das Bestandsobjekt nie verkauft; im "
                            "Verkaufsjahr der Betrag, der abgelöst wird.",
    "prg_ErgebnisHalten": "Bestandsobjekt: Miete + weitere Einnahmen − Erhaltung − weitere "
                          "Ausgaben − AfA, als würde es nie verkauft (Baseline); Neuobjekt 0.",
    "prg_Wirkung": "Ergebnis − Zinsen Darlehen im Plan minus dasselbe bei Halten. Nicht "
                   "verkauftes Bestandsobjekt 0, verkauftes ab dem Folgejahr das entfallende "
                   "Ergebnis (negativ), Neuobjekt sein Ergebnis nach Kreditzinsen. Grundlage "
                   "der Ergebnisbrücke.",
    "prg_VerkaufNr": "Zeile im Blatt Verkäufe, deren aktiver Verkauf das Objekt enthält (auch "
                     "als weitere Kostenstelle eines Pakets); 0 ohne Verkauf.",
    "prg_Unternehmen": "1 = Kostenstelle der Art Unternehmen (Blatt Objekte), keine Immobilie.",
    # Darlehen
    "dl_Teil": "Hauptkredit (oberer Teil des Blatts) oder Zweitkredit (unterer Teil).",
    "dl_Betrag": "Hauptkredit: Kredit − davon Zweitkredit aus dem Blatt Neuobjekte; "
                 "Zweitkredit: davon Zweitkredit. Ohne Betrag bleibt die Zeile leer.",
    "dl_Tilgungsfrei": "Jahre nach dem Kauf mit Zins, aber ohne Tilgung; zählen zur Laufzeit. "
                       "Zweitkredit leer = wie Hauptkredit.",
    "dl_Rate": "Annuität: Kredit × (Zinssatz + anfängliche Tilgung), gleich bleibend. Linear: "
               "Kredit × Tilgungssatz je Jahr. Endfällig: keine laufende Tilgung.",
    "dl_UeJahr": "Restschuld Zweitkredit nach Laufzeit = auf Hauptkredit: Kaufjahr + Laufzeit "
                 "des Zweitkredits; zum Ende dieses Jahrs geht seine Restschuld auf den "
                 "Hauptkredit über. In beiden Zeilen, leer ohne Übernahme.",
    "dl_UeBetrag": "Zweitkredit: Betrag − Tilgung bis einschließlich Übernahmejahr (die "
                   "reguläre Rate des letzten Jahrs ist gezahlt); Hauptkredit: derselbe Betrag, "
                   "er erhöht dort die Restschuld. Leer, wenn das Jahr nach dem Raster liegt.",
    "dl_RateNeu": "Hauptkredit nach der Übernahme: Rate + übernommener Betrag × (Zinssatz + "
                  "Tilgung), linear Tilgung + Betrag × Tilgungssatz, endfällig 0. Gilt ab dem "
                  "Jahr nach der Übernahme.",
    "dl_Getilgt": "Jahr, in dem die Restschuld 0 erreicht (Zweitkredit mit Übernahme: das "
                  "Übernahmejahr); „nach …“: am Rasterende nicht getilgt.",
}

# Liquidität: je Szenario dieselbe Tabelle
LIQUIDITAET = {
    "einnahmen": "Summe Miete + weitere Einnahmen aller Objekte dieses Jahres (Prognose); "
                 "Baseline: Bestand ohne Verkäufe.",
    "ausgaben": "Summe Erhaltung + weitere Ausgaben aller Objekte dieses Jahres.",
    "afa": "Summe AfA. A: nach § 6b-Kürzung der Neuobjekte; B und C: von den vollen AK.",
    "ergebnis": "Einnahmen − Ausgaben − AfA.",
    "verkauf": "A: steuerwirksam laut Rücklagenspiegel (sofort versteuerte Gewinne + Auflösung + "
               "Zuschlag). B und C: jeder Veräußerungsgewinn sofort.",
    "zins": "Liquidität kumuliert am Vorjahresende × Rendite Alternativanlage.",
    "kreditzins": "Summe der Zinsen aller Darlehen: Bestandsobjekte bis zum Verkauf (Blatt "
                  "Objekte), Kredite der Neuobjekte (Blatt Darlehen); voll abziehbar.",
    "zve": "laufendes Ergebnis + steuerwirksam aus Verkauf + Zinsertrag − Zinsen Darlehen.",
    "vortrag_genutzt": "MIN(Verlustvortrag Vorjahr, positives Ergebnis).",
    "bemessung": "positives Ergebnis − genutzter Verlustvortrag.",
    "vortrag": "Vorjahr − genutzt + neuer Verlust.",
    "steuer": "Bemessungsgrundlage × angewandter Steuersatz; KG mit Verlustausgleich + "
              "(Ergebnis − Bemessungsgrundlage) × Satz auf Verluste.",
    "verkaufserloes": "Nettoerlöse der gültigen Verkäufe dieses Jahres.",
    "rueckfluss": "Nettoerlös − Veräußerungsgewinn = steuerneutral zurückfließender Buchwert.",
    "kauf": "Kaufpreis + Nebenkosten der Neuobjekte mit diesem Kaufjahr (B: ohne Neuobjekte mit "
            "Rücklage).",
    "kredit": "Kreditauszahlungen der Neuobjekte mit diesem Kaufjahr.",
    "tilgung": "Summe der Tilgungen; im Verkaufsjahr eines Bestandsobjekts mit der Ablösung "
               "seiner Restschuld. Mindert nur die Liquidität, nicht das Ergebnis.",
    "abloesung": "Teil der Tilgung, der im Verkaufsjahr die Restschuld des verkauften "
                 "Objekts ablöst; der Rest ist planmäßige Tilgung.",
    "zufluss": "Einnahmen − Ausgaben + Zinsertrag − Zinsen + Verkaufserlöse − Steuer − Kauf "
               "+ Kreditauszahlung − Tilgung. AfA fließt nicht ab.",
    "kum": "Vorjahr + freier Mittelzufluss; negativ = Finanzierungslücke.",
    "restschuld": "Summe der Restschulden aller Darlehen am Jahresende.",
}

# Auswertung: je Szenario dieselbe Tabelle
AUSWERTUNG = {
    "ergebnis": "aus der Liquidität: Einnahmen − Ausgaben − AfA.",
    "verkauf": "aus der Liquidität: steuerwirksam aus Verkauf und Rücklage.",
    "zins": "aus der Liquidität: Zinsertrag Alternativanlage.",
    "kreditzins": "aus der Liquidität: Zinsen Darlehen (Bestand und Neuobjekte).",
    "guv": "laufendes Ergebnis + steuerwirksam aus Verkauf + Zinsertrag − Zinsen Darlehen.",
    "steuer": "aus der Liquidität.",
    "nach_steuer": "Gesamt-GuV − Steuer.",
    "steuer_kum": "Summe der Steuer bis zu diesem Jahr.",
    "verkehrswert": "Summe der Verkehrswerte der Objekte im Bestand am Jahresende.",
    "buchwert": "Buchwert Gebäude + G+B der Objekte im Bestand (B und C ohne § 6b-Kürzung).",
    "stille_reserven": "Verkehrswert − Buchwert.",
    "ruecklage": "Bestand der § 6b-Rücklage (nur A).",
    "vortrag": "Verlustvortrag am Jahresende aus der Liquidität.",
    "liquiditaet": "Liquidität kumuliert aus der Liquidität.",
    "restschuld": "Restschuld aller Darlehen am Jahresende.",
    "vermoegen": "Verkehrswert + Liquidität − Restschuld.",
    "latente_steuer": "(stille Reserven + Rücklage − Verlustvortrag) × angewandter "
                      "Steuersatz, mindestens 0: Steuer bei Verkauf aller Objekte zum "
                      "Verkehrswert. KG mit Verlustausgleich: der Teil bis zum Vortrag × "
                      "Satz auf Verluste.",
    "vermoegen_netto": "Gesamtvermögen − latente Steuer: die Vergleichsgröße der Szenarien.",
}

SPALTEN.update({
    # Anlagen
    "anl_Gruppe": "G+B, abnutzbar (Gebäude, BGA, im Bau) oder nicht im Modell.",
    "anl_Zugang": "Jahr des AHK-Datums bei G+B und Gebäude; das früheste je Objekt ist das "
                  "Kaufjahr.",
    "anl_BWBasis": "Buchwert Stand, fortgeschrieben mit der AfA bis zum Ende des Basisjahrs.",
    "anl_AfABasis": "AfA im Basisjahr nach Methode und AfA p. a.",
    "anl_Status": "OK oder was fehlt; nur Zeilen mit OK zählen für das Objekt.",
})
# Hilfsspalten je Quelle im Blatt Neuobjekte (eingeklappt)
for _n in (1, 2, 3):
    SPALTEN.update({
        f"ne_Q{_n}RLGeb": f"Rücklage Gebäude der Quelle {_n}, abzüglich was Zeilen darüber "
                          "schon übertragen haben; 0, wenn die Quelle unbekannt ist oder das "
                          "Kaufjahr nicht zwischen Bildungs- und Fristjahr liegt.",
        f"ne_Q{_n}RLGuB": f"Rücklage G+B der Quelle {_n}, abzüglich bereits übertragener ü2 und "
                          "ü3.",
        f"ne_Q{_n}Ue1": f"ü1 der Quelle {_n}: MIN(Rücklage Gebäude, AK Gebäude neu − ü1 der "
                        "Quellen davor).",
        f"ne_Q{_n}Ue2": f"ü2 der Quelle {_n}: MIN(Rücklage G+B, AK G+B neu − ü2 der Quellen "
                        "davor).",
        f"ne_Q{_n}Ue3": f"ü3 der Quelle {_n}: MIN(Rest Rücklage G+B, AK Gebäude neu − alle ü1 − "
                        "ü3 der Quellen davor).",
        f"ne_Q{_n}Erloes": f"Nettoerlös des Verkaufs hinter Quelle {_n}, der in diesen Kauf "
                           "fließt: höchstens der noch offene Teil von Kaufpreis + Nebenkosten "
                           "und was Zeilen darüber nicht eingesetzt haben; 0 ohne "
                           "Verkaufserlös zur Finanzierung = ja.",
    })

# Ergebnisbrücke: Tooltips der Zeilen
BRUECKE = {
    "halten": "Gesamt-GuV vor Steuern der Baseline (Blatt Auswertung): alle Kostenstellen "
              "werden gehalten, keine Verkäufe, keine Käufe.",
    "halten_immobilien": "Laufendes Ergebnis der Immobilien nach Darlehenszinsen bei Halten.",
    "halten_unternehmen": "Kostenstellen der Art Unternehmen (Blatt Objekte): Einnahmen − "
                          "Ausgaben − AfA − Zinsen, aus dem Basisjahr fortgeschrieben. Sie "
                          "ändern sich durch die Entscheidungen nicht, solange sie nicht "
                          "verkauft werden.",
    "halten_zins": "Zinsertrag auf die Liquidität bei Halten (Rendite Alternativanlage).",
    "paket": "Zeile im Blatt Verkäufe, aus deren Rücklage oder Erlös der Kauf finanziert ist "
             "(erste Quelle im Blatt Neuobjekte, die einen Verkauf nennt); 0 = Kauf ohne "
             "Verkauf. Darüber ordnet die Brücke den Kauf seinem Paket zu.",
    "entscheidungen": "Summe der Pakete und der Käufe ohne Verkauf: die Wirkung aller "
                      "Entscheidungen vor Steuern, ohne den Zinseffekt der Liquidität.",
    "ohne": "Käufe, die keine Rücklage und keinen Erlös eines Verkaufs nutzen.",
    "verkaeufe": "Je Verkauf: steuerwirksamer Gewinn im Verkaufsjahr, bei § 6b minus "
                 "Einstellung in die Rücklage, im Fristjahr plus Auflösung und Zuschlag; "
                 "dazu ab dem Folgejahr das entfallende laufende Ergebnis (Miete − Kosten − "
                 "AfA − Darlehenszinsen) der verkauften Kostenstellen.",
    "vk_veraeusserung": "Gewinn und Rücklage aller Verkäufe, gleich der Spalte steuerwirksam "
                        "im Rücklagenspiegel.",
    "vk_laufend": "Entfallendes laufendes Ergebnis aller verkauften Kostenstellen.",
    "neuobjekte": "Alle Käufe, mit oder ohne Paket. Je Kauf: Miete − Kosten − AfA − Kreditzinsen. Die AfA rechnet von den um "
                  "die übertragene Rücklage geminderten AK, mit Sonder-AfA § 7b.",
    "zins": "Zinsertrag Plan minus Halten: Verkaufserlöse, Kaufpreise und Steuer verändern die "
            "Liquidität und damit deren Verzinsung.",
    "plan": "Gesamt-GuV vor Steuern des Plans A (Blatt Auswertung).",
    "kontrolle": "Halten + Summe Pakete und Käufe ohne Verkauf + Zinseffekt − Plan; muss 0 sein, "
                 "sonst meldet das Prüfungsblatt einen Fehler.",
    "nach_wirkung": "Ergebnis nach Steuern Plan A minus Halten: Wirkung aller Entscheidungen "
                    "zusammen, nach Steuer mit Verlustvortrag.",
}


# Cashflow: Tooltips der Zeilen; Abflüsse negativ
CASHFLOW = {
    "jahresueberschuss": "Ergebnis nach Steuern des gewählten Szenarios (Blatt Auswertung).",
    "afa": "AfA mindert das Ergebnis, fließt aber nicht ab (Blatt Liquidität).",
    "gewinne": "Veräußerungsgewinne stehen im Ergebnis, sind aber kein Geldzufluss: das Geld "
               "kommt als Verkaufserlös im Investitions-Cashflow. Baseline: keine Verkäufe.",
    "bildung": "Nur A: Einstellung in die § 6b-Rücklage neutralisiert den Gewinn, das Geld "
               "ist noch da (wie die Zunahme einer Rückstellung).",
    "aufloesung": "Nur A: Auflösung der Rücklage im Fristjahr erhöht das Ergebnis ohne "
                  "Einzahlung.",
    "zuschlag": "Nur A: Gewinnzuschlag § 6b erhöht das Ergebnis ohne Einzahlung.",
    "operativ": "Geld aus dem laufenden Geschäft nach Zinsen und Steuer: Jahresüberschuss "
                "plus Aufwand ohne Auszahlung minus Ertrag ohne Einzahlung. Gleich der Summe "
                "der davon-Zeilen darunter.",
    "einnahmen": "Mieten und weitere Einnahmen aller Objekte (Blatt Liquidität).",
    "ausgaben": "Erhaltung und weitere Ausgaben aller Objekte.",
    "zinsertrag": "Zins auf die Liquidität am Vorjahresende (Rendite Alternativanlage).",
    "kreditzins": "Zinsen aller Darlehen: Bestandsobjekte und Kredite der Neuobjekte.",
    "steuer": "Steuer des Jahres mit Verlustvortrag (Blatt Liquidität).",
    "verkaufserloes": "Nettoerlöse der gültigen Verkäufe dieses Jahres.",
    "kauf": "Kaufpreis und Nebenkosten der Neuobjekte mit diesem Kaufjahr. Die Übertragung "
            "der § 6b-Rücklage zahlt nichts und erscheint hier nicht.",
    "investition": "Verkaufserlöse minus Käufe.",
    "kredit": "Kreditauszahlungen der Neuobjekte zum Kauf.",
    "tilgung": "Tilgung nach Tilgungsplan bzw. Rate − Zins, ohne Ablösung bei Verkauf.",
    "abloesung": "Restschuld eines verkauften Bestandsobjekts, abgelöst im Verkaufsjahr.",
    "finanzierung": "Kreditauszahlung minus Tilgung und Ablösung.",
    "veraenderung": "operativ + Investition + Finanzierung = freier Mittelzufluss im Blatt "
                    "Liquidität.",
    "anfang": "Liquidität Ende des Vorjahres; das Raster beginnt mit 0.",
    "ende": "Liquidität Anfang + Veränderung.",
    "kontrolle": "Liquidität Ende minus Liquidität kumuliert im Blatt Liquidität; muss 0 sein, "
                 "sonst meldet das Blatt Prüfung einen Fehler.",
    "free_cashflow": "Operativer Cashflow plus Investitions-Cashflow: was nach Käufen und "
                     "Verkäufen für Tilgung und Anlage bleibt.",
    "kapitaldienst": "Zinsen und planmäßige Tilgung der Darlehen, ohne Ablösung bei Verkauf.",
    "nach_kapitaldienst": "Operativer Cashflow minus planmäßige Tilgung: was nach Bank und "
                          "Finanzamt aus dem laufenden Geschäft übrig bleibt.",
    "dscr": "Kapitaldienstdeckung (DSCR) = (Einnahmen − Ausgaben) / Kapitaldienst. Unter 1 "
            "(rot) decken die Mieten Zins und Tilgung nicht; leer ohne Kapitaldienst.",
}
