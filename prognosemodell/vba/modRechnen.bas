Attribute VB_Name = "modRechnen"
Option Explicit

' Neuberechnung: in Excel vollstaendig; wo CalculateFull fehlt (LibreOffice im
' Pruefskript), die normale Neuberechnung.
Public Sub NeuBerechnen()
    On Error Resume Next
    Application.CalculateFull
    If Err.Number <> 0 Then
        Err.Clear
        Application.Calculate
    End If
    On Error GoTo 0
End Sub

Public Sub NeuBerechnenStarten()
    NeuBerechnen
    MsgBox Txt("Neu berechnet." & vbLf & "Plausibilit{ae}tspr{ue}fung: ") & _
        Bereich("pr_Gesamt").Value, vbInformation, "Prognosemodell"
End Sub
