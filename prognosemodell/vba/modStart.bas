Attribute VB_Name = "modStart"
Option Explicit

' Steuerung des Prognosemodells (Projektplan Abschnitt 5 und 19, Etappe 9).
' VBA rechnet nicht: Die Makros schreiben nur Eingabewerte und loesen die
' Neuberechnung aus, alle Ergebnisse entstehen in den Zellformeln.
' Der Quelltext ist bewusst ohne Umlaute; Meldungstexte laufen ueber Txt().

Private Const KNOPF_PREFIX As String = "knopf_"

' Schaltflaechen auf dem Parameterblatt (untereinander) und dem Blatt Varianten
' (nebeneinander) anlegen, falls sie fehlen (beim Oeffnen)
Public Sub Einrichten()
    Dim ws As Worksheet
    Set ws = Bereich("par_Basisjahr").Worksheet
    KnoepfeAnlegen ws, KNOPF_PREFIX, Schaltflaechen(), ws.Range("F5"), False
    Set ws = Bereich("var_Tabelle").Worksheet
    KnoepfeAnlegen ws, KNOPF_PREFIX & "var_", Array( _
        Array("Variante festhalten", "VarianteFesthaltenStarten"), _
        Array("Markierte Variante einspielen", "VarianteEinspielenStarten")), _
        ws.Range("E1"), True
End Sub

Private Sub KnoepfeAnlegen(ByVal ws As Worksheet, ByVal prefix As String, _
        ByVal knoepfe As Variant, ByVal start As Range, ByVal nebeneinander As Boolean)
    Dim i As Long, links As Double, oben As Double
    If KnopfVorhanden(ws, prefix & "0") Then Exit Sub
    For i = LBound(knoepfe) To UBound(knoepfe)
        links = start.Left
        oben = start.Top
        If nebeneinander Then links = links + i * 230 Else oben = oben + i * 28
        With ws.Buttons.Add(links, oben + 1, IIf(nebeneinander, 220, 260), 24)
            .Name = prefix & i
            .Caption = Txt(knoepfe(i)(0))
            .OnAction = knoepfe(i)(1)
            ' frei schwebend: eingeklappte Zeilen verschieben die Knoepfe nicht
            On Error Resume Next
            .Placement = 3
            On Error GoTo 0
        End With
    Next i
End Sub

Private Function Schaltflaechen() As Variant
    Schaltflaechen = Array( _
        Array("Plausibilit{ae}t pr{ue}fen", "PruefungStarten"), _
        Array("Neu berechnen", "NeuBerechnenStarten"), _
        Array("Variante festhalten (Blatt Varianten)", "VarianteFesthaltenStarten"), _
        Array("Variante einspielen (Blatt Varianten)", "VarianteEinspielenStarten"), _
        Array("Objekt anlegen", "ObjektAnlegenStarten"), _
        Array("Objekt duplizieren", "ObjektDuplizierenStarten"), _
        Array("Objekt entfernen", "ObjektEntfernenStarten"), _
        Array("Annahmen wiederherstellen (leere Felder)", "AnnahmenWiederherstellenStarten"), _
        Array("Objekte -> Kostenstellen ({ue}berschriebene Werte)", _
              "InKostenstelleUebernehmenStarten"), _
        Array("Leere Prognosebl{oe}cke aus-/einblenden", "LeereBloeckeUmschaltenStarten"))
End Function

Private Function KnopfVorhanden(ByVal ws As Worksheet, ByVal knopfName As String) As Boolean
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = knopfName Then
            KnopfVorhanden = True
            Exit Function
        End If
    Next s
End Function

' Benannter Bereich der Mappe (par_*, obj_*, vg_*, pr_* ...)
Public Function Bereich(ByVal bereichsname As String) As Range
    Set Bereich = ThisWorkbook.Names(bereichsname).RefersToRange
End Function

' Platzhalter in Meldungstexten durch Sonderzeichen ersetzen
Public Function Txt(ByVal s As String) As String
    s = Replace(s, "{ae}", ChrW(228))
    s = Replace(s, "{oe}", ChrW(246))
    s = Replace(s, "{ue}", ChrW(252))
    s = Replace(s, "{Ae}", ChrW(196))
    s = Replace(s, "{Oe}", ChrW(214))
    s = Replace(s, "{Ue}", ChrW(220))
    s = Replace(s, "{ss}", ChrW(223))
    s = Replace(s, "{sect}", ChrW(167))
    s = Replace(s, "{euro}", ChrW(8364))
    Txt = s
End Function

' Betrag fuer Meldungen, z. B. 1.234.567 EUR
Public Function Euro(ByVal betrag As Variant) As String
    Euro = Format(betrag, "#,##0") & " " & Txt("{euro}")
End Function
