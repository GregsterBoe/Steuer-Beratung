Attribute VB_Name = "modStart"
Option Explicit

' Steuerung des Prognosemodells (Projektplan Abschnitt 5 und 19, Etappe 9).
' VBA rechnet nicht: Die Makros schreiben nur Eingabewerte und loesen die
' Neuberechnung aus, alle Ergebnisse entstehen in den Zellformeln.
' Der Quelltext ist bewusst ohne Umlaute; Meldungstexte laufen ueber Txt().

Private Const KNOPF_PREFIX As String = "knopf_"

' Schaltflaechen auf dem Parameterblatt anlegen, falls sie fehlen (beim Oeffnen)
Public Sub Einrichten()
    Dim ws As Worksheet, knoepfe As Variant, i As Long
    Dim links As Double, oben As Double
    Set ws = Bereich("par_Basisjahr").Worksheet
    If KnopfVorhanden(ws, KNOPF_PREFIX & "0") Then Exit Sub
    knoepfe = Schaltflaechen()
    links = ws.Range("F5").Left
    oben = ws.Range("F5").Top
    For i = LBound(knoepfe) To UBound(knoepfe)
        With ws.Buttons.Add(links, oben + i * 28, 260, 24)
            .Name = KNOPF_PREFIX & i
            .Caption = Txt(knoepfe(i)(0))
            .OnAction = knoepfe(i)(1)
        End With
    Next i
End Sub

Private Function Schaltflaechen() As Variant
    Schaltflaechen = Array( _
        Array("Plausibilit{ae}t pr{ue}fen", "PruefungStarten"), _
        Array("Neu berechnen", "NeuBerechnenStarten"), _
        Array("Variante festhalten (Blatt Varianten)", "VarianteFesthaltenStarten"), _
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
