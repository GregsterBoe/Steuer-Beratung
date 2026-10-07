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
    "**Daten bereitstellen:** Kostenstellen-Datei (DATEV-BWA, ein Blatt je Kostenstelle, in "
    "Excel gespeichert) und, falls vorhanden, das Anlagenverzeichnis (DATEV „Inventarübersicht“). "
    "Die Mappe wird daraus erzeugt; das übernimmt die technische Betreuung.",
    "**Blatt Start öffnen:** Handlungsempfehlung, Belastbarkeit der Daten und Farblegende prüfen.",
    "**Annahmen prüfen:** Blatt Parameter (Steuersatz, Steigerungen, Alternativrendite) und "
    "blaue Zellen im Blatt Objekte. Wo bessere Werte bekannt sind, eintippen.",
    "**Planung erfassen:** geplante Verkäufe im Blatt Verkäufe, Käufe im Blatt Neuobjekte.",
    "**Ergebnis lesen:** Start, Übersicht und Vergleich; zuvor im Blatt Prüfung sicherstellen, "
    "dass keine Fehler offen sind.",
], "List Number")

# ---------------------------------------------------------------- 2
doc.add_heading("2. Die wichtigsten Blätter", level=1)
tabelle(["Blatt", "Wofür"], [
    ["**Start**", "Empfehlung (Option mit dem höchsten Endvermögen nach latenter Steuer), "
                  "Vorsprung gegenüber Halten, tiefster Liquiditätsstand, Anleitung mit Links"],
    ["**Parameter**", "Alle Sätze und Annahmen zentral; Schaltflächen der Makros"],
    ["**Objekte**", "Ein Bestandsobjekt je Zeile: Miete, Kosten, AK, Kaufjahr, Verkehrswert, Status"],
    ["**Verkäufe**", "ObjektID, Jahr, Preis, Kosten, Anteil G+B lt. Kaufvertrag, § 6b ja/nein, "
                     "„reinvestieren = ja“ legt ein Neuobjekt aus Annahmen an"],
    ["**Neuobjekte**", "Kaufjahr, Kaufpreis, Anteil G+B, AfA-Satz und -Methode, Quelle-Rücklage, "
                       "optional „Kostenstelle Neukauf“"],
    ["**Anlagen**", "Anlagenverzeichnis: AK, Buchwert und AfA je Anlage und Jahr"],
    ["**Übersicht / Vergleich**", "Wert der Immobilien und Gesamtvermögen je Jahr; Szenarien "
                                  "nebeneinander mit Differenzen"],
    ["**Prüfung**", "Alle Plausibilitätsprüfungen als Fehler, Warnung oder Hinweis"],
    ["**Alle Objekte, KSt-Blätter**", "Ist und Plan im DATEV-BWA-Format je Kostenstelle und "
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

# ---------------------------------------------------------------- 4
doc.add_heading("4. Kostenstellenblätter sind führend", level=1)
punkte([
    "Die Blätter je Kostenstelle und „Alle Objekte“ sind die Datenbasis. Werte des Basisjahrs "
    "(gelbe Spalte) werden **dort** geändert; sie fließen automatisch in Objekte, Prognose und "
    "„Alle Objekte“.",
    "Im Blatt Objekte verweisen Miete, Erhaltung, weitere Einnahmen und Ausgaben sowie die AfA "
    "lt. Buchhaltung auf das Kostenstellenblatt (grün).",
    "Wird ein solcher Wert im Blatt Objekte überschrieben, rechnet nur die Prognose damit. Die "
    "Zelle wird orange, das Prüfungsblatt warnt. Die Schaltfläche **„Objekte -> Kostenstellen“** "
    "überträgt den Wert ins Kostenstellenblatt und stellt die Verknüpfung wieder her.",
    "Die Monatsspalten des Basisjahrs sind eingeklappt; das „+“ über der Jahresspalte öffnet sie.",
])

doc.add_heading("Neukauf-Kostenstellen (KSt 31–35)", level=2)
absatz("Alle Kostenstellen hinter „KSt 9999“ gelten als geplante Käufe. Sie erhalten ein eigenes "
       "Blatt (Basisjahr und Planjahre gelb), zählen aber nicht zu „Alle Objekte“. Trägt man die "
       "Kostenstelle im Blatt Neuobjekte unter „Kostenstelle Neukauf“ ein, übernimmt die Prognose "
       "ab dem Jahr nach dem Kauf Miete, Einnahmen, Erhaltung und Ausgaben aus diesem Blatt; "
       "Jahre ohne Wert werden mit der Steigerung fortgeschrieben. Die AfA rechnet weiter das "
       "Modell aus Kaufpreis, Anteil G+B, AfA-Satz und § 6b-Übertragung.")

# ---------------------------------------------------------------- 5
doc.add_heading("5. Steuerliche Logik in Kürze", level=1)
punkte([
    "**Verkauf** zum Jahresende; Miete und AfA laufen im Verkaufsjahr noch. Der Erlös wird nach "
    "Kaufvertrag, sonst nach Verkehrswertanteil auf Gebäude und G+B aufgeteilt.",
    "**§ 6b-Rücklage** aus den positiven Teilgewinnen, getrennt nach Gebäude und G+B. Frist vier "
    "Jahre, mit begonnenem Neubau sechs. Nicht übertragene Rücklage wird im Fristjahr aufgelöst, "
    "mit 6 % Zuschlag je Jahr.",
    "**Übertragung** in fester Reihenfolge: Gebäudegewinn auf Gebäude, G+B-Gewinn auf G+B, Rest "
    "auf das Gebäude. Die AfA des Neuobjekts läuft von der geminderten Basis.",
    "**Steuer** = (laufendes Ergebnis + steuerwirksamer Rücklagenbetrag) × Grenzsteuersatz; "
    "Verluste werden vorgetragen. Freie Liquidität wird mit der Alternativrendite verzinst.",
    "**Erhaltung** steigt ab 30 Jahren Gebäudealter zusätzlich; ab 50 Jahren fällt eine "
    "Großmaßnahme an (Standard 15 % des Gebäudewerts). Neuobjekte tragen anfangs die Hälfte.",
])

doc.add_heading("Szenarien", level=2)
tabelle(["Szenario", "Inhalt"], [
    ["A Plan", "§ 6b-Kette wie erfasst"],
    ["B", "Gewinn sofort versteuert, kein Neukauf aus der Rücklage, Geld in der Alternativanlage"],
    ["C", "sofort versteuert, Neuobjekte trotzdem gekauft, volle AfA-Basis"],
    ["Baseline", "alles halten"],
], [3.0, 13.6])
absatz("A − C zeigt die reine Wirkung von § 6b: der Zins auf die gestundete Steuer.")

# ---------------------------------------------------------------- 6
doc.add_heading("6. Prüfung und Schaltflächen", level=1)
punkte([
    "**Fehler:** Status ungleich OK (Objekte, Verkäufe, Neuobjekte), Fristverstoß, Steuerwelt "
    "nicht GmbH. Vor jeder Auswertung beheben.",
    "**Warnung:** Vorbesitzzeit für § 6b zu kurz, Rücklage nicht voll übertragen, Drei-Objekt-"
    "Grenze, Abweichung vom Kostenstellenblatt, unbekannte Neukauf-Kostenstelle.",
    "**Hinweis:** Frist nach Prognoseende, negative Liquidität, fehlender Verkehrswert, Anlagen "
    "ohne Objekt.",
])
absatz("In der Makro-Fassung (.xlsm) stehen auf dem Parameterblatt: Plausibilität prüfen, Neu "
       "berechnen, Variante festhalten (Kennzahlen als feste Werte im Blatt Varianten, etwa für "
       "verschiedene Verkaufsjahre), Objekt anlegen, duplizieren, entfernen, Annahmen "
       "wiederherstellen, Objekte -> Kostenstellen, leere Prognoseblöcke aus-/einblenden. "
       "Makros rechnen nichts; alle Ergebnisse stehen in den Zellformeln.")

# ---------------------------------------------------------------- 7
doc.add_heading("7. Praxistipps", level=1)
punkte([
    "Bei verkauften Objekten AK, Kaufjahr und Anteil G+B belegen; orange Zellen bestimmen den "
    "Gewinn direkt.",
    "Der **Schnellcheck** (schlanke Mappe) genügt für eine erste Einschätzung: ObjektID und Miete "
    "je Objekt plus ein geplanter Verkauf.",
    "Zinsen und Tilgung sind noch nicht enthalten (Betrachtung vor Finanzierung).",
    "**Alle Steuersätze, Fristen und Annahmen vor dem Echteinsatz durch den zuständigen "
    "Berufsträger prüfen.**",
])

doc.save(sys.argv[1])
