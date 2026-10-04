Attribute VB_Name = "modSzenario"
Option Explicit

' Szenarien A, B, C rechnen und im Blatt Vergleich speichern (Projektplan Abschnitt 15).
' Das Makro setzt nur par_Szenario und kopiert die Spalte vg_Aktuell als Werte
' in vg_A, vg_B bzw. vg_C; gerechnet wird in den Formeln.

Public Sub SzenarioSetzen(ByVal szenario As String)
    Bereich("par_Szenario").Value = szenario
    NeuBerechnen
End Sub

Public Sub SzenarioSpeichern(ByVal szenario As String)
    Dim ws As Worksheet
    Set ws = Bereich("vg_Aktuell").Worksheet
    ws.Unprotect
    Bereich("vg_" & szenario).Value = Bereich("vg_Aktuell").Value
    ws.Protect
End Sub

' Alle drei Szenarien rechnen und speichern, danach das vorher aktive wieder einstellen
Public Sub SzenarienVergleichen()
    Dim vorher As String, sz As Variant
    vorher = CStr(Bereich("par_Szenario").Value)
    For Each sz In Array("A", "B", "C")
        SzenarioSetzen CStr(sz)
        SzenarioSpeichern CStr(sz)
    Next sz
    SzenarioSetzen vorher
End Sub

Public Sub VergleichLeeren()
    Dim ws As Worksheet, sz As Variant
    Set ws = Bereich("vg_Aktuell").Worksheet
    ws.Unprotect
    For Each sz In Array("A", "B", "C")
        Bereich("vg_" & sz).ClearContents
    Next sz
    ws.Protect
End Sub

Public Sub SzenarienVergleichenStarten()
    Dim v As Range
    If Not PruefungBestanden("Szenariovergleich") Then Exit Sub
    Application.ScreenUpdating = False
    SzenarienVergleichen
    Application.ScreenUpdating = True
    Bereich("vg_Aktuell").Worksheet.Activate
    Set v = Bereich("vg_Vermoegen")
    MsgBox Txt("Szenarien A, B und C gerechnet und im Blatt Vergleich gespeichert." & vbLf & _
        vbLf & "Endverm{oe}gen nach Steuern:") & vbLf & _
        "A  " & Euro(v.Cells(1, 2).Value) & vbLf & _
        "B  " & Euro(v.Cells(1, 3).Value) & vbLf & _
        "C  " & Euro(v.Cells(1, 4).Value) & vbLf & vbLf & _
        "Aktives Szenario wieder " & Bereich("par_Szenario").Value & ".", _
        vbInformation, "Prognosemodell"
End Sub

Public Sub SzenarioAStarten()
    SzenarioRechnenStarten "A"
End Sub

Public Sub SzenarioBStarten()
    SzenarioRechnenStarten "B"
End Sub

Public Sub SzenarioCStarten()
    SzenarioRechnenStarten "C"
End Sub

Private Sub SzenarioRechnenStarten(ByVal szenario As String)
    SzenarioSetzen szenario
    MsgBox "Szenario " & szenario & Txt(" ist aktiv." & vbLf & _
        "Endverm{oe}gen nach Steuern: ") & Euro(Bereich("vg_Vermoegen").Cells(1, 1).Value) & _
        vbLf & Txt("Plausibilit{ae}tspr{ue}fung: ") & Bereich("pr_Gesamt").Value, _
        vbInformation, "Prognosemodell"
End Sub
