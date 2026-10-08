"""Kompakte Bedienungsanleitung für Steuerberater als .docx.

Aufruf: python docs/bedienungsanleitung.py docs/Bedienungsanleitung_Steuerberatung.docx
"""
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

doc = Document()
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2)
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(10.5)
st.paragraph_format.space_after = Pt(4)
for name, size in (("Title", 22), ("Heading 1", 14), ("Heading 2", 11.5)):
    doc.styles[name].font.name = "Calibri"
    doc.styles[name].font.size = Pt(size)
    doc.styles[name].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)


def absatz(text):
    p = doc.add_paragraph()
    _fett(p, text)
    return p


def _fett(p, text):
    """**fett** im Text auswerten."""
    teile = text.split("**")
    for i, t in enumerate(teile):
        if t:
            p.add_run(t).bold = i % 2 == 1


def punkte(liste, stil="List Bullet"):
    for t in liste:
        p = doc.add_paragraph(style=stil)
        p.paragraph_format.space_after = Pt(2)
        _fett(p, t)


def schattieren(zelle, farbe):
    tc = zelle._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), farbe)
    tc.append(shd)


def tabelle(kopf, zeilen, breiten, farben=None):
    t = doc.add_table(rows=1, cols=len(kopf))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, k in enumerate(kopf):
        z = t.rows[0].cells[i]
        z.text = ""
        z.paragraphs[0].add_run(k).bold = True
        schattieren(z, "D9E2F3")
    for n, zeile in enumerate(zeilen):
        zellen = t.add_row().cells
        for i, w in enumerate(zeile):
            zellen[i].text = ""
            _fett(zellen[i].paragraphs[0], w)
        if farben and farben[n]:
            schattieren(zellen[0], farben[n])
    t.autofit = False
    for i, b in enumerate(breiten):
        t.columns[i].width = Cm(b)
    for zeile in t.rows:
        for i, b in enumerate(breiten):
            zeile.cells[i].width = Cm(b)
            for p in zeile.cells[i].paragraphs:
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.keep_with_next = True   # Tabelle nicht trennen
                for r in p.runs:
                    r.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


# ---------------------------------------------------------------- Titel
doc.add_paragraph("Prognosemodell V+V", style="Title")
absatz("**Kurzanleitung für die Steuerberatung:** 20-Jahres-Prognose für Immobilien im "
       "Betriebsvermögen einer GmbH mit AfA, Verkauf, § 6b-Rücklage und Reinvestition. "
       "Die Mappe vergleicht Halten, § 6b-Kette und sofortiges Versteuern.")

# ---------------------------------------------------------------- 1
doc.add_heading("1. Ablauf in fünf Schritten", level=1)
punkte([
    "**Blatt Start öffnen:** Handlungsempfehlung, Belastbarkeit der Daten und Farblegende ansehen.",
    "**Kostenstellen prüfen:** Werte des Basisjahrs in den Kostenstellenblättern kontrollieren "
    "und dort korrigieren.",
    "**Annahmen prüfen:** im Blatt Parameter zuerst die Kernparameter (★, fett: Steuersatz, "
    "Steigerungen, Alternativrendite), dann die blauen Zellen im Blatt Objekte. Wo bessere "
    "Werte bekannt sind, eintippen.",
    "**Planung erfassen:** geplante Verkäufe im Blatt Verkäufe, Käufe im Blatt Neuobjekte.",
    "**Prüfen und lesen:** „Plausibilität prüfen“ drücken, Fehler beheben, dann Start, "
    "Übersicht und Vergleich lesen.",
], "List Number")

# ---------------------------------------------------------------- 2
doc.add_heading("2. Die wichtigsten Blätter", level=1)
tabelle(["Blatt", "Wofür"], [
    ["**Start**", "Empfehlung (Option mit dem höchsten Endvermögen nach latenter Steuer), "
                  "Vorsprung gegenüber Halten, tiefster Liquiditätsstand, Anleitung mit Links"],
    ["**Parameter**", "Alle Sätze und Annahmen zentral; Schaltflächen (Abschnitt 7)"],
    ["**Objekte**", "Ein Bestandsobjekt je Zeile: Miete, Kosten, AK, Kaufjahr, Verkehrswert, Status"],
    ["**Verkäufe**", "ObjektID, Jahr, Preis, Kosten, Anteil G+B lt. Kaufvertrag, § 6b ja/nein, "
                     "„reinvestieren = ja“ legt ein Neuobjekt aus Annahmen an"],
    ["**Neuobjekte**", "Kaufjahr, Kaufpreis, Anteil G+B, AfA-Satz und -Methode, bis zu drei "
                       "Quell-Rücklagen, „Kostenstelle Neukauf“, Finanzierung (Abschnitt 4)"],
    ["**Darlehen**", "Tilgungsplan je Kredit: Zins, Tilgung und Restschuld je Jahr"],
    ["**Anlagen**", "Anlagenverzeichnis: AK, Buchwert und AfA je Anlage und Jahr"],
    ["**Übersicht / Vergleich**", "Ergebnis je Jahr und Szenarien nebeneinander (Abschnitt 5)"],
    ["**Prüfung**", "Alle Plausibilitätsprüfungen als Fehler, Warnung oder Hinweis"],
    ["**Alle Objekte, KSt-Blätter**", "Ist und Plan im BWA-Format je Kostenstelle und "
                                      "als Summe; Verkauf und Kauf als eigenes Blatt"],
], [4.2, 12.4])

# ---------------------------------------------------------------- 3
doc.add_heading("3. Farben der Eingabezellen", level=1)
tabelle(["Farbe", "Bedeutung", "Was tun?"], [
    ["rot", "Pflichtangabe fehlt oder Annahme gelöscht", "Wert eintragen"],
    ["orange", "Annahme bei verkauftem Objekt (bestimmt Gewinn und Rücklage) oder Wert "
               "weicht vom Kostenstellenblatt ab", "Wert belegen bzw. abgleichen"],
    ["blau", "Annahme aus dem Parameterblatt", "bei besserer Kenntnis überschreiben"],
    ["grün", "aus Buchhaltung, Kostenstellenblatt oder Anlagen", "nichts; zählt nicht als Annahme"],
    ["gelb", "händische Eingabe", "—"],
    ["grau", "Formel", "nicht überschreiben"],
], [2.2, 8.4, 6.0], ["F4B6B6", "F8CBAD", "BDD7EE", "C6E0B4", "FFF2CC", "D9D9D9"])
absatz("Ein eingetippter Wert ersetzt die Annahme. Wird ein Feld geleert, stellt die Schaltfläche "
       "„Annahmen wiederherstellen“ die Annahme wieder her.")
absatz("**Die Mappe erklärt sich selbst:** Spaltenköpfe in Objekte, Verkäufe, Neuobjekte und "
       "Anlagen sind rot mit * (Pflichtfeld), blau (optional) oder grau (berechnet). Ein rotes "
       "Dreieck am Kopf heißt: Maus darüber zeigt die Erläuterung, bei berechneten Spalten die "
       "Herleitung in Worten; das gilt auch für Prognose, Rücklagen, Darlehen, Liquidität und "
       "Auswertung. Im Blatt Parameter zeigt die Maus über der Bezeichnung, was der Parameter "
       "beeinflusst; Abschnitte ohne Kernparameter sind eingeklappt (+ am linken Rand).")

# ---------------------------------------------------------------- 4
doc.add_heading("4. Kostenstellenblätter sind führend", level=1)
punkte([
    "Die Blätter je Kostenstelle und „Alle Objekte“ sind die Datenbasis. Werte des Basisjahrs "
    "(gelbe Spalte) werden **dort** geändert; sie fließen automatisch in Objekte, Prognose und "
    "„Alle Objekte“.",
    "Im Blatt Objekte verweisen Miete, Erhaltung, weitere Einnahmen und Ausgaben, die AfA "
    "lt. Buchhaltung und der Zinsaufwand (1310) auf das Kostenstellenblatt (grün).",
    "Wird ein solcher Wert im Blatt Objekte überschrieben, rechnet nur die Prognose damit. Die "
    "Zelle wird orange, das Prüfungsblatt warnt. Die Schaltfläche **„Objekte -> Kostenstellen“** "
    "überträgt den Wert ins Kostenstellenblatt und stellt die Verknüpfung wieder her.",
    "Die Monatsspalten des Basisjahrs sind eingeklappt; das „+“ über der Jahresspalte öffnet sie.",
])

doc.add_heading("Neukauf-Kostenstellen (KSt 31–35)", level=2)
absatz("Alle Kostenstellen hinter „KSt 9999“ gelten als geplante Käufe. Sie erhalten ein eigenes "
       "Blatt (Basisjahr und Planjahre gelb). Wählt man die Kostenstelle im Blatt Neuobjekte unter "
       "„Kostenstelle Neukauf“ aus, übernimmt die Prognose ab dem Jahr nach dem Kauf Miete, "
       "Einnahmen, Erhaltung und Ausgaben aus diesem Blatt; Jahre ohne Wert werden mit der "
       "Steigerung fortgeschrieben. Das Blatt ist dann zugleich das Blatt des Neuobjekts: AfA und "
       "Kreditzinsen kommen grau aus dem Modell, das Neuobjekt zählt in „Alle Objekte“. Kaufpreis, "
       "Anteil G+B und AfA-Satz bleiben Eingaben im Blatt Neuobjekte.")

doc.add_heading("Rücklagen und Finanzierung eines Kaufs", level=2)
punkte([
    "**Quell-Rücklagen:** bis zu drei Rücklagen je Neuobjekt aus der Auswahlliste. Übertragen "
    "werden erst alle Gebäudegewinne auf das Gebäude, dann die G+B-Gewinne auf G+B, der Rest auf "
    "das Gebäude. Ob die Übertragung im Einzelfall gewollt ist, entscheidet der Berater.",
    "**Finanzierungsbedarf** = Kaufpreis + Nebenkosten − Nettoerlös der Quell-Verkäufe.",
    "**Finanzierung Rest:** leer bzw. „Eigenmittel“ zahlt den Bedarf aus der Liquidität; "
    "„Kredit“ finanziert ihn. Kreditbetrag leer = ganzer Bedarf, sonst der eingetragene Betrag, "
    "der Rest kommt aus Eigenmitteln.",
    "**Kreditangaben:** Zinssatz, Tilgungsart (Annuität mit anfänglicher Tilgung, linear oder "
    "endfällig), Tilgung p. a., optional die Laufzeit, nach der die Restschuld getilgt wird.",
    "Zinsen mindern das Ergebnis und die Steuer, die Tilgung nur die Liquidität, die Restschuld "
    "das Gesamtvermögen. Ist der Kredit getilgt, entfallen Zins und Tilgung (Blatt Darlehen).",
])

doc.add_heading("Darlehen der Bestandsobjekte", level=2)
punkte([
    "**Grob, ohne weitere Daten:** Der Zinsaufwand des Basisjahrs (BWA 1310) läuft bis zum "
    "Verkaufsjahr weiter und sinkt je Jahr um den Satz vom Parameterblatt (Standard 3 %). "
    "Tilgung und Ablösung beim Verkauf fehlen; die Prüfung weist darauf hin.",
    "**Genau:** Im Blatt Objekte die **Restschuld zum 31.12.** des Basisjahrs eintragen "
    "(Saldenliste oder Bankauszug), möglichst auch Zinssatz und Jahresrate. Dann rechnet die "
    "Prognose Zins und Tilgung je Jahr und löst die Restschuld im Verkaufsjahr aus dem Erlös "
    "ab. Fehlen Zinssatz oder Rate, stehen Annahmen (blau) in den Zellen.",
    "Eine Vorfälligkeitsentschädigung gehört in die Verkaufskosten.",
])

# ---------------------------------------------------------------- 5
doc.add_heading("5. Ergebnis lesen", level=1)
punkte([
    "**Start:** empfohlene Option mit dem höchsten Endvermögen nach latenter Steuer, Vorsprung "
    "gegenüber Halten und tiefster Liquiditätsstand.",
    "**Übersicht:** Wert der Immobilien und Gesamtvermögen je Jahr, als Tabelle und Diagramm, "
    "mit und ohne die geplanten Verkäufe.",
    "**Vergleich:** Kennzahlen der Szenarien am Ende des Prognosezeitraums nebeneinander.",
    "**Varianten:** mit „Variante festhalten“ gesicherte Stände, etwa verschiedene Verkaufsjahre.",
])
tabelle(["Szenario", "Inhalt"], [
    ["A Plan", "§ 6b-Kette wie erfasst"],
    ["B", "Gewinn sofort versteuert, kein Neukauf aus der Rücklage, Geld angelegt"],
    ["C", "sofort versteuert, Neuobjekte trotzdem gekauft"],
    ["Baseline", "alles halten"],
], [3.0, 13.6])

# ---------------------------------------------------------------- 6
doc.add_heading("6. Prüfung", level=1)
punkte([
    "**Fehler:** Status ungleich OK in Objekten, Verkäufen oder Neuobjekten, Fristverstoß. "
    "Vor jeder Auswertung beheben.",
    "**Warnung:** Vorbesitzzeit für § 6b zu kurz, Rücklage nicht voll übertragen, Drei-Objekt-"
    "Grenze, Abweichung vom Kostenstellenblatt, unbekannte Neukauf-Kostenstelle, Verkauf "
    "ohne Restschuld trotz Zinsaufwand.",
    "**Hinweis:** Frist nach Prognoseende, negative Liquidität, fehlender Verkehrswert, Anlagen "
    "ohne Objekt.",
])

# ---------------------------------------------------------------- 7
doc.add_heading("7. Schaltflächen (Blatt Parameter)", level=1)
tabelle(["Schaltfläche", "Wirkung"], [
    ["Plausibilität prüfen", "zeigt alle Fehler, Warnungen und Hinweise"],
    ["Neu berechnen", "rechnet die Mappe vollständig neu"],
    ["Variante festhalten", "speichert die Kennzahlen des Vergleichs im Blatt Varianten"],
    ["Objekt anlegen / duplizieren / entfernen", "pflegt die Zeilen im Blatt Objekte"],
    ["Annahmen wiederherstellen", "setzt geleerte Felder wieder auf die Annahme"],
    ["Objekte -> Kostenstellen", "überträgt überschriebene Werte ins Kostenstellenblatt"],
    ["Leere Prognoseblöcke aus-/einblenden", "blendet ungenutzte Zeilen der Prognose aus"],
], [6.2, 10.4])

# ---------------------------------------------------------------- 8
doc.add_heading("8. Praxistipps", level=1)
punkte([
    "Bei verkauften Objekten AK, Kaufjahr und Anteil G+B belegen; orange Zellen bestimmen den "
    "Gewinn direkt.",
    "Graue Formelzellen nicht überschreiben; Änderungen immer in gelben oder blauen Zellen.",
    "Bei Objekten mit Darlehen die Restschuld eintragen, vor allem vor einem Verkauf: sonst "
    "fehlt die Ablösung, Liquidität und Kapitalanlage sind zu hoch.",
    "**Alle Steuersätze, Fristen und Annahmen vor dem Echteinsatz durch den zuständigen "
    "Berufsträger prüfen.**",
])

doc.save(sys.argv[1])
