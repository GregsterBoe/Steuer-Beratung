Attribute VB_Name = "modVarianten"
Option Explicit

' Varianten der Eingaben vergleichen (Projektplan Abschnitt 15 und 19): Das Makro
' haelt die Kennzahlen des Blatts Vergleich als feste Werte im Blatt Varianten fest.
' Danach Eingaben aendern (Verkaufsjahr, Preis, Parameter) und erneut festhalten.

' Erste freie Zeile im Blatt Varianten; 0 = Tabelle voll
Public Function FreieVariante() As Long
    Dim tabelle As Range, i As Long
    Set tabelle = Bereich("var_Tabelle")
    For i = 1 To tabelle.Rows.Count
        If CStr(tabelle.Cells(i, 1).Value) = "" Then
            FreieVariante = i
            Exit Function
        End If
    Next i
End Function

' Kennzahlen festhalten; liefert die Zeile in der Tabelle (1 = erste Variante)
Public Function VarianteFesthalten(ByVal bezeichnung As String) As Long
    Dim tabelle As Range, pos As Long
    bezeichnung = Trim$(bezeichnung)
    If bezeichnung = "" Then Err.Raise vbObjectError + 520, "modVarianten", "Bezeichnung fehlt."
    pos = FreieVariante()
    If pos = 0 Then
        Err.Raise vbObjectError + 521, "modVarianten", "Blatt Varianten ist voll."
    End If
    NeuBerechnen
    Set tabelle = Bereich("var_Tabelle")
    tabelle.Cells(pos, 1).Value = bezeichnung
    tabelle.Cells(pos, 2).Value = Now
    ' erste Kennzahl im Blatt Vergleich: Endvermoegen nach latenter Steuer
    tabelle.Cells(pos, 3).Value = Bereich("vg_A").Cells(1, 1).Value
    tabelle.Cells(pos, 4).Value = Bereich("vg_B").Cells(1, 1).Value
    tabelle.Cells(pos, 5).Value = Bereich("vg_C").Cells(1, 1).Value
    tabelle.Cells(pos, 6).Value = Bereich("vg_Baseline").Cells(1, 1).Value
    tabelle.Cells(pos, 7).Value = Bereich("vg_DiffB").Cells(1, 1).Value
    tabelle.Cells(pos, 8).Value = Bereich("vg_DiffC").Cells(1, 1).Value
    tabelle.Cells(pos, 9).Value = Application.WorksheetFunction.Sum(Bereich("liq_Steuer"))
    tabelle.Cells(pos, 10).Value = Bereich("pr_Gesamt").Value
    VarianteFesthalten = pos
End Function

Public Sub VariantenLeeren()
    Bereich("var_Tabelle").ClearContents
End Sub

Public Sub VarianteFesthaltenStarten()
    Dim bezeichnung As String, pos As Long
    If Not PruefungBestanden("Festhalten") Then Exit Sub
    bezeichnung = InputBox("Bezeichnung der Variante:", "Variante festhalten", _
        "Variante " & FreieVariante())
    If bezeichnung = "" Then Exit Sub
    On Error GoTo Fehler
    pos = VarianteFesthalten(bezeichnung)
    Bereich("var_Tabelle").Worksheet.Activate
    Bereich("var_Tabelle").Cells(pos, 1).Select
    MsgBox Txt("Variante festgehalten. Endverm{oe}gen A: ") & _
        Euro(Bereich("var_Tabelle").Cells(pos, 3).Value), vbInformation, "Prognosemodell"
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub
