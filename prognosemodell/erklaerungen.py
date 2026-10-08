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
    "par_Basisjahr", "par_Prognosejahre", "par_Steuersatz", "par_Mietsteig",
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
    "par_Steuerwelt": "Nur GmbH ist gerechnet; jeder andere Wert ergibt einen Fehler in der "
                      "Prüfung.",
    "par_Steuersatz": "Steuer je Jahr = Bemessungsgrundlage × Satz (Liquidität) und latente "
                      "Steuer = stille Reserven + Rücklage − Verlustvortrag, × Satz "
                      "(Auswertung). Wirkt damit auf Liquidität, Zinsertrag und das "
                      "Endvermögen aller Szenarien.",
    "par_Mietsteig": "Miete und weitere Einnahmen je Jahr = Basiswert × (1 + Satz)^Jahre. "
                     "Gilt für Bestand, Neuobjekte und Neukauf-Kostenstellen ohne Planwert.",
    "par_Erhaltsteig": "Erhaltung je Jahr wächst mit diesem Satz; auch Großmaßnahmen "
                       "(heutige Preise bis zum Jahr der Maßnahme).",
    "par_Kostensteig": "Weitere Ausgaben je Jahr = Basiswert × (1 + Satz)^Jahre.",
    "par_Wertsteig": "Verkehrswert je Jahr = Verkehrswert aktuell × (1 + Satz)^Jahre. Bestimmt "
                     "die Verkaufspreis-Annahme, stille Reserven, latente Steuer und das "
                     "Endvermögen. Neuobjekte: ab dem Kaufpreis im Kaufjahr.",
    "par_GrESt": "Zur Information; die Nebenkosten der Neuobjekte kommen aus der Eingabe bzw. "
                 "der Annahme „Kaufnebenkosten in % des Kaufpreises“.",
    "par_AfADegressiv": "Neuobjekte mit AfA-Methode degressiv: AfA = Satz × Restbuchwert, "
                        "Wechsel zur linearen AfA, sobald diese höher ist.",
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
    "par_AnnGebaeudealter": "Nur ohne Baujahr: Baujahr = Basisjahr − Alter. Steuert Alterung "
                            "der Erhaltung und Jahr der Großmaßnahme.",
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
    "par_ErhAlterungAb": "Erhaltung aller Objekte: ab diesem Gebäudealter kommt je Jahr die "
                         "zusätzliche Steigerung dazu.",
    "par_ErhAlterung": "Zusätzliche Steigerung der Erhaltung je Jahr über dem Alter.",
    "par_NeuErhAnlaufJahre": "Neuobjekte: so viele Jahre nach dem Kauf gilt der Anlauffaktor.",
    "par_NeuErhAnlaufFaktor": "Neuobjekte: Erhaltung in den Anlaufjahren = Erhaltung × Faktor.",
    "par_SanAlter": "Nur ohne Eingabe im Blatt Objekte: Jahr der Großmaßnahme = Baujahr + Alter.",
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
    "vk_Vorbesitz": "Verkaufsjahr − Kaufjahr (Blatt Objekte). Unter der Mindest-Vorbesitzzeit "
                    "(Parameter) ist § 6b nicht möglich.",
    "vk_BuchwertGeb": "Buchwert Gebäude am Ende des Verkaufsjahrs aus dem Blatt Prognose (nach "
                      "der AfA des Verkaufsjahrs).",
    "vk_AKGuB": "AK G+B aus dem Blatt Objekte = Buchwert G+B (keine AfA).",
    "vk_Nettoerloes": "Verkaufspreis − Verkaufskosten.",
    "vk_AnteilGuB": "Anteil G+B lt. Kaufvertrag; leer: 1 − Verkehrswertanteil Gebäude aus dem "
                    "Blatt Objekte.",
    "vk_ErloesGeb": "Nettoerlös − Erlösanteil G+B.",
    "vk_ErloesGuB": "Nettoerlös × Anteil G+B verwendet.",
    "vk_GewinnGeb": "Erlösanteil Gebäude − Buchwert Gebäude Ende Verkaufsjahr. Positiv: "
                    "Rücklage Gebäude bei § 6b; negativ: Verlust wirkt sofort.",
    "vk_GewinnGuB": "Erlösanteil G+B − AK G+B. Positiv: Rücklage G+B bei § 6b.",
    "vk_Gewinn": "Gewinn Gebäude + Gewinn G+B. Ohne § 6b sofort steuerpflichtig im "
                 "Verkaufsjahr; mit § 6b gehen die positiven Teile in die Rücklage.",
    "vk_Status": "OK oder der erste Fehler (ObjektID unbekannt, doppelt verkauft, Jahr "
                 "fehlt/außerhalb, Preis fehlt, Aufteilung fehlt). „§ 6b unzulässig“: die Zeile "
                 "rechnet, der Gewinn wird aber sofort versteuert. Andere Fehler: die Zeile "
                 "rechnet nicht mit.",
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
    "ne_Gueltig": "1 = rechnet im Modell: Pflichtfelder da, NeuID eindeutig und keine ObjektID, "
                  "Kaufjahr im Raster.",
    "ne_AKGuBNeu": "(Kaufpreis + Nebenkosten) × Anteil G+B.",
    "ne_AKGebNeu": "(Kaufpreis + Nebenkosten) × (1 − Anteil G+B).",
    "ne_RLGeb": "Summe der Quellen: Rücklage Gebäude, soweit nicht schon von Zeilen darüber "
                "verwendet.",
    "ne_RLGuB": "Summe der Quellen: Rücklage G+B, soweit nicht schon verwendet.",
    "ne_Ue1": "Gebäudegewinn auf das Gebäude: höchstens AK Gebäude neu.",
    "ne_Ue2": "G+B-Gewinn auf G+B: höchstens AK G+B neu.",
    "ne_Ue3": "Rest des G+B-Gewinns auf das Gebäude: höchstens AK Gebäude neu − ü1.",
    "ne_UeGesamt": "ü1 + ü2 + ü3: mindert die AK und geht als Übertragung in den "
                   "Rücklagenspiegel.",
    "ne_AfABasis": "AK Gebäude neu − ü1 − ü3: Basis für AfA und Buchwert Gebäude.",
    "ne_AKGuB": "AK G+B neu − ü2: Buchwert G+B.",
    "ne_MitQuelle": "1 = mindestens eine Quelle angegeben; das Neuobjekt entfällt in Szenario B.",
    "ne_AKGesamt": "Kaufpreis + Kaufnebenkosten.",
    "ne_Erloes": "Nettoerlös der Quell-Verkäufe, der in diesen Kauf fließt (höchstens Kaufpreis "
                 "+ Nebenkosten, soweit nicht schon von Zeilen darüber eingesetzt).",
    "ne_Bedarf": "Kaufpreis + Nebenkosten − Einsatz Verkaufserlös: was anders finanziert werden "
                 "muss.",
    "ne_Kredit": "Bei Finanzierung Rest = Kredit: Kreditbetrag, leer = ganzer "
                 "Finanzierungsbedarf. Auszahlung zum Ende des Kaufjahrs, Plan im Blatt "
                 "Darlehen.",
    "ne_Eigen": "Finanzierungsbedarf − Kredit: kommt aus der Liquidität.",
    "ne_Status": "OK oder Fehler. Pflichtfeld fehlt, NeuID doppelt bzw. wie Bestandsobjekt, "
                 "Kaufjahr außerhalb Raster: das Neuobjekt rechnet nicht. Meldungen zu Quellen "
                 "(unbekannt, Kauf vor Bildung, nach Fristjahr) und Kredit: es rechnet, aber "
                 "ohne diese Übertragung bzw. mit dem gemeldeten Mangel.",
    # Prognose
    "prg_Aktiv": "1 = Objekt erzielt in diesem Jahr Miete und AfA: Bestand bis einschließlich "
                 "Verkaufsjahr, Neuobjekt ab dem Jahr nach dem Kauf.",
    "prg_Miete": "Bestand: Miete Basisjahr × (1 + Mietsteigerung)^Jahre. Neuobjekt: Kaufpreis × "
                 "Mietrendite, bzw. Planwert der Neukauf-Kostenstelle.",
    "prg_Erhaltung": "Erhaltung Basisjahr, fortgeschrieben mit Erhaltungssteigerung, Alterung "
                     "über dem Gebäudealter, plus Großmaßnahme im Jahr der Maßnahme.",
    "prg_AfA": "AfA laut Blatt Anlagen bzw. AfA-Plan, sonst AfA je Jahr bis der Restbuchwert "
               "verbraucht ist. Neuobjekt: AfA-Basis × Satz (bzw. degressiv).",
    "prg_Verkehrswert": "Verkehrswert aktuell × (1 + Wertsteigerung)^Jahre, unabhängig vom "
                        "Verkauf. Neuobjekt: Kaufpreis × (1 + Wertsteigerung)^Jahre seit Kauf.",
    "prg_Buchwert": "Buchwert Vorjahr − AfA, nicht unter 0.",
    "prg_Ergebnis": "Miete + weitere Einnahmen − Erhaltung − weitere Ausgaben − AfA.",
    "prg_Bestand": "1 = am Jahresende im Bestand (Verkaufsjahr schon 0).",
    "prg_KreditZins": "Neuobjekt: Zinsen des Kredits (Blatt Darlehen). Bestand mit Restschuld: "
                      "Restschuld Vorjahresende × Zinssatz; ohne Restschuld: Zinsaufwand "
                      "Basisjahr × (1 + Veränderung)^Jahre bis zum Verkauf.",
    "prg_Tilgung": "Neuobjekt: Blatt Darlehen. Bestand: Rate − Zins, höchstens die Restschuld; "
                   "im Verkaufsjahr die ganze Restschuld (Ablösung aus dem Erlös).",
    "prg_Restschuld": "Restschuld Vorjahresende − Tilgung.",
    "prg_ZinsHalten": "Zinsen, als würde das Bestandsobjekt nie verkauft (Baseline).",
    "prg_TilgungHalten": "Tilgung, als würde das Bestandsobjekt nie verkauft (Baseline).",
    "prg_RestschuldHalten": "Restschuld, als würde das Bestandsobjekt nie verkauft; im "
                            "Verkaufsjahr der Betrag, der abgelöst wird.",
    # Darlehen
    "dl_Betrag": "Kredit aus dem Blatt Neuobjekte; 0 = keine Finanzierung per Kredit.",
    "dl_Rate": "Annuität: Kredit × (Zinssatz + anfängliche Tilgung), gleich bleibend. Linear: "
               "Kredit × Tilgungssatz je Jahr. Endfällig: keine laufende Tilgung.",
    "dl_Getilgt": "Jahr, in dem die Restschuld 0 erreicht; „nach …“: am Rasterende nicht "
                  "getilgt.",
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
    "steuer": "Bemessungsgrundlage × Grenzsteuersatz.",
    "verkaufserloes": "Nettoerlöse der gültigen Verkäufe dieses Jahres.",
    "rueckfluss": "Nettoerlös − Veräußerungsgewinn = steuerneutral zurückfließender Buchwert.",
    "kauf": "Kaufpreis + Nebenkosten der Neuobjekte mit diesem Kaufjahr (B: ohne Neuobjekte mit "
            "Rücklage).",
    "kredit": "Kreditauszahlungen der Neuobjekte mit diesem Kaufjahr.",
    "tilgung": "Summe der Tilgungen; im Verkaufsjahr eines Bestandsobjekts mit der Ablösung "
               "seiner Restschuld. Mindert nur die Liquidität, nicht das Ergebnis.",
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
    "latente_steuer": "(stille Reserven + Rücklage − Verlustvortrag) × Grenzsteuersatz, "
                      "mindestens 0: Steuer bei Verkauf aller Objekte zum Verkehrswert.",
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
                           "und was Zeilen darüber nicht eingesetzt haben.",
    })
