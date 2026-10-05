Attribute VB_Name = "modPruefung"
Option Explicit

' Plausibilitaetspruefungen: Gerechnet wird im Blatt Pruefung (Formeln), damit
' Fehler auch ohne Makros sichtbar bleiben. Das Makro rechnet neu, liest das
' Ergebnis und meldet es.

Public Function AnzahlFehler() As Long
    NeuBerechnen
    AnzahlFehler = CLng(Bereich("pr_Fehler").Value)
End Function

Public Function AnzahlWarnungen() As Long
    NeuBerechnen
    AnzahlWarnungen = CLng(Bereich("pr_Warnungen").Value)
End Function

' Je auffaellige Pruefung eine Zeile: Ergebnis, Text, Anzahl betroffener Zeilen
Public Function Auffaelligkeiten() As String
    Dim ergebnis As Range, i As Long, liste As String
    Set ergebnis = Bereich("pr_Ergebnis")
    For i = 1 To ergebnis.Rows.Count
        If CStr(ergebnis.Cells(i, 1).Value) <> "OK" Then
            liste = liste & ergebnis.Cells(i, 1).Value & ": " & _
                Bereich("pr_Bezeichnung").Cells(i, 1).Value & " (" & _
                Bereich("pr_Anzahl").Cells(i, 1).Value & ")" & vbLf
        End If
    Next i
    Auffaelligkeiten = liste
End Function

Public Sub PruefungStarten()
    Dim liste As String, symbol As Long
    NeuBerechnen
    liste = Auffaelligkeiten()
    Bereich("pr_Ergebnis").Worksheet.Activate
    If liste = "" Then
        MsgBox Txt("Plausibilit{ae}tspr{ue}fung: alles OK."), vbInformation, "Prognosemodell"
    Else
        symbol = vbInformation
        If Bereich("pr_Fehler").Value > 0 Then symbol = vbExclamation
        MsgBox Txt("Plausibilit{ae}tspr{ue}fung: ") & Bereich("pr_Gesamt").Value & vbLf & _
            vbLf & liste, symbol, "Prognosemodell"
    End If
End Sub

' Vor Makros, die rechnen: bei Fehlern nachfragen, ob trotzdem weiter
Public Function PruefungBestanden(ByVal aktion As String) As Boolean
    If AnzahlFehler() = 0 Then
        PruefungBestanden = True
    Else
        PruefungBestanden = (MsgBox(Txt("Die Plausibilit{ae}tspr{ue}fung meldet ") & _
            Bereich("pr_Gesamt").Value & "." & vbLf & vbLf & Auffaelligkeiten() & vbLf & _
            aktion & " trotzdem starten?", vbYesNo + vbExclamation, "Prognosemodell") = vbYes)
    End If
End Function
