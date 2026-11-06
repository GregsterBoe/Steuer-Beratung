Attribute VB_Name = "modVarianten"
Option Explicit

' Varianten der Eingaben vergleichen (Projektplan Abschnitt 15 und 19): Das Makro
' haelt die Kennzahlen des Blatts Vergleich, den tiefsten Liquiditaetsstand, die
' Verkaeufe und Kaeufe und die wichtigsten Parameter als feste Werte im Blatt
' Varianten fest. Zusaetzlich sichert es alle Eingaben (Parameter, Objekte, Verkaeufe,
' Neuobjekte, Neukauf-KSt) im ausgeblendeten Blatt Variantenspeicher; "Variante
' einspielen" schreibt sie zurueck. Woher eine Spalte ihren Wert nimmt, steht in der
' ausgeblendeten Zeile var_Quellen.

' Vor jedem gespeicherten Eintrag, damit Formeln und Zahlen als Text im Speicher stehen
Private Const PRAEFIX As String = "~"
Private Const JA As String = "ja"

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
    Dim tabelle As Range, quellen As Range, pos As Long, j As Long
    bezeichnung = Trim$(bezeichnung)
    If bezeichnung = "" Then Err.Raise vbObjectError + 520, "modVarianten", "Bezeichnung fehlt."
    pos = FreieVariante()
    If pos = 0 Then
        Err.Raise vbObjectError + 521, "modVarianten", "Blatt Varianten ist voll."
    End If
    NeuBerechnen
    Set tabelle = Bereich("var_Tabelle")
    Set quellen = Bereich("var_Quellen")
    EingabenSpeichern pos
    For j = 1 To quellen.Columns.Count
        tabelle.Cells(pos, j).Value = Kennzahl(CStr(quellen.Cells(1, j).Value), bezeichnung)
    Next j
    VarianteFesthalten = pos
End Function

' Wert einer Spalte laut Quelle: benannter Bereich (erste Zelle) oder "*Angabe"
Private Function Kennzahl(ByVal quelle As String, ByVal bezeichnung As String) As Variant
    Dim art As String, rest As String, p As Long
    If Left$(quelle, 1) <> "*" Then
        Kennzahl = Bereich(quelle).Cells(1, 1).Value
        Exit Function
    End If
    art = Mid$(quelle, 2)
    p = InStr(art, ":")
    If p > 0 Then
        rest = Mid$(art, p + 1)
        art = Left$(art, p - 1)
    End If
    Select Case art
        Case "Bezeichnung": Kennzahl = bezeichnung
        Case "Datum": Kennzahl = Now
        Case "Summe": Kennzahl = Application.WorksheetFunction.Sum(Bereich(rest))
        Case "Minimum": Kennzahl = TiefsterStand(rest, False)
        Case "MinimumJahr": Kennzahl = TiefsterStand(rest, True)
        Case "Verkaeufe": Kennzahl = Vorgaenge("vk_ID", "vk_Jahr", "vk_Preis", "vk_Aktiv")
        Case "Kaeufe": Kennzahl = Vorgaenge("ne_ID", "ne_Kaufjahr", "ne_Kaufpreis", "ne_Aktiv")
        Case "Gespeichert": Kennzahl = JA
        Case Else: Kennzahl = ""
    End Select
End Function

' Tiefster Wert des Bereichs (Liquiditaet je Jahr) bzw. das Jahr dazu (liq_Jahr)
Private Function TiefsterStand(ByVal bereichsname As String, ByVal jahr As Boolean) As Variant
    Dim werte As Range, i As Long, best As Long
    Set werte = Bereich(bereichsname)
    For i = 1 To werte.Cells.Count
        If IsNumeric(werte.Cells(i).Value) And CStr(werte.Cells(i).Value) <> "" Then
            If best = 0 Then
                best = i
            ElseIf CDbl(werte.Cells(i).Value) < CDbl(werte.Cells(best).Value) Then
                best = i
            End If
        End If
    Next i
    If best = 0 Then
        TiefsterStand = ""
    ElseIf jahr Then
        TiefsterStand = Bereich("liq_Jahr").Cells(best).Value
    Else
        TiefsterStand = werte.Cells(best).Value
    End If
End Function

' Aktive Vorgaenge als Text: "ID Jahr Preis; ..." (Preis in Tsd. Euro)
Private Function Vorgaenge(ByVal idName As String, ByVal jahrName As String, _
        ByVal preisName As String, ByVal aktivName As String) As String
    Dim ids As Range, i As Long, s As String, preis As Variant
    Set ids = Bereich(idName)
    For i = 1 To ids.Rows.Count
        If Trim$(CStr(ids.Cells(i, 1).Value)) <> "" And _
                LCase$(Trim$(CStr(Bereich(aktivName).Cells(i, 1).Value))) <> "nein" Then
            preis = Bereich(preisName).Cells(i, 1).Value
            If s <> "" Then s = s & "; "
            s = s & Trim$(CStr(ids.Cells(i, 1).Value)) & " " & _
                CStr(Bereich(jahrName).Cells(i, 1).Value)
            If IsNumeric(preis) And CStr(preis) <> "" Then
                s = s & " " & Format(CDbl(preis) / 1000, "#,##0") & Txt(" T{euro}")
            End If
        End If
    Next i
    Vorgaenge = s
End Function

' Zeilen eines Variantenblocks im Speicher: Summe der Zeilen aller Bereiche
Private Function BlockHoehe() As Long
    Dim namen As Range, i As Long
    Set namen = Bereich("var_Bereiche")
    For i = 1 To namen.Rows.Count
        BlockHoehe = BlockHoehe + Bereich(CStr(namen.Cells(i, 1).Value)).Rows.Count
    Next i
End Function

Private Function BlockBreite() As Long
    Dim namen As Range, i As Long, n As Long
    Set namen = Bereich("var_Bereiche")
    For i = 1 To namen.Rows.Count
        n = Bereich(CStr(namen.Cells(i, 1).Value)).Columns.Count
        If n > BlockBreite Then BlockBreite = n
    Next i
End Function

' Erste Zelle des Blocks der Variante pos im Speicher
Private Function BlockStart(ByVal pos As Long) As Range
    Set BlockStart = Bereich("var_Speicher").Offset((pos - 1) * BlockHoehe(), 0)
End Function

' Alle Eingaben (FormulaR1C1: Annahmeformeln bleiben Formeln) in den Block pos schreiben
Private Sub EingabenSpeichern(ByVal pos As Long)
    Dim namen As Range, ziel As Range, b As Range, werte As Variant
    Dim i As Long, r As Long, c As Long
    Set namen = Bereich("var_Bereiche")
    Set ziel = BlockStart(pos)
    ziel.Resize(BlockHoehe(), BlockBreite()).ClearContents
    For i = 1 To namen.Rows.Count
        Set b = Bereich(CStr(namen.Cells(i, 1).Value))
        werte = Formeln(b)
        For r = 1 To b.Rows.Count
            For c = 1 To b.Columns.Count
                If CStr(werte(r, c)) <> "" Then werte(r, c) = PRAEFIX & CStr(werte(r, c))
            Next c
        Next r
        ziel.Resize(b.Rows.Count, b.Columns.Count).Value = werte
        Set ziel = ziel.Offset(b.Rows.Count, 0)
    Next i
End Sub

' FormulaR1C1 eines Bereichs, immer als zweidimensionales Feld
Private Function Formeln(ByVal b As Range) As Variant
    Dim werte As Variant, feld() As Variant
    werte = b.FormulaR1C1
    If IsArray(werte) Then
        Formeln = werte
    Else
        ReDim feld(1 To 1, 1 To 1)
        feld(1, 1) = werte
        Formeln = feld
    End If
End Function

Private Function Gespeichert(ByVal pos As Long) As Boolean
    Gespeichert = (CStr(Bereich("var_Gespeichert").Cells(pos, 1).Value) = JA)
End Function

' Eingaben der Variante pos (Zeile im Blatt Varianten) zurueckschreiben; nur Zellen, die
' sich unterscheiden. Liefert die Anzahl geaenderter Zellen.
Public Function VarianteEinspielen(ByVal pos As Long) As Long
    Dim namen As Range, quelle As Range, b As Range, alt As Variant, neu As Variant
    Dim i As Long, r As Long, c As Long, soll As String, n As Long
    If pos < 1 Or pos > Bereich("var_Tabelle").Rows.Count Then
        Err.Raise vbObjectError + 522, "modVarianten", "Variante " & pos & " gibt es nicht."
    End If
    If CStr(Bereich("var_Bezeichnung").Cells(pos, 1).Value) = "" Then
        Err.Raise vbObjectError + 522, "modVarianten", "Variante " & pos & " gibt es nicht."
    End If
    If Not Gespeichert(pos) Then
        Err.Raise vbObjectError + 523, "modVarianten", _
            Txt("F{ue}r diese Variante sind keine Eingaben gespeichert.")
    End If
    Set namen = Bereich("var_Bereiche")
    Set quelle = BlockStart(pos)
    For i = 1 To namen.Rows.Count
        Set b = Bereich(CStr(namen.Cells(i, 1).Value))
        alt = Formeln(b)
        neu = quelle.Resize(b.Rows.Count, b.Columns.Count).Value
        If Not IsArray(neu) Then neu = Formeln(quelle)
        For r = 1 To b.Rows.Count
            For c = 1 To b.Columns.Count
                soll = CStr(neu(r, c))
                If Left$(soll, 1) = PRAEFIX Then soll = Mid$(soll, 2) Else soll = ""
                If CStr(alt(r, c)) <> soll Then
                    b.Cells(r, c).FormulaR1C1 = soll
                    n = n + 1
                End If
            Next c
        Next r
        Set quelle = quelle.Offset(b.Rows.Count, 0)
    Next i
    NeuBerechnen
    VarianteEinspielen = n
End Function

' Position einer Variante ueber die Bezeichnung; 0 = nicht gefunden
Public Function VariantePosition(ByVal bezeichnung As String) As Long
    Dim namen As Range, i As Long
    Set namen = Bereich("var_Bezeichnung")
    For i = 1 To namen.Rows.Count
        If Trim$(CStr(namen.Cells(i, 1).Value)) = Trim$(bezeichnung) And _
                CStr(namen.Cells(i, 1).Value) <> "" Then
            VariantePosition = i
            Exit Function
        End If
    Next i
End Function

Public Sub VariantenLeeren()
    Bereich("var_Tabelle").ClearContents
    Bereich("var_Speicher").Resize(Bereich("var_Tabelle").Rows.Count * BlockHoehe(), _
        BlockBreite()).ClearContents
End Sub

' ---- Bedienung ueber Schaltflaechen ----

Public Sub VarianteFesthaltenStarten()
    Dim bezeichnung As String, pos As Long
    If Not PruefungBestanden("Festhalten") Then Exit Sub
    bezeichnung = InputBox("Bezeichnung der Variante:", "Variante festhalten", _
        "Variante " & FreieVariante())
    If bezeichnung = "" Then Exit Sub
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    pos = VarianteFesthalten(bezeichnung)
    Application.ScreenUpdating = True
    Bereich("var_Tabelle").Worksheet.Activate
    Bereich("var_Tabelle").Cells(pos, 1).Select
    MsgBox Txt("Variante festgehalten. Endverm{oe}gen A: ") & _
        Euro(Bereich("var_A").Cells(pos, 1).Value) & vbLf & _
        Txt("Liquidit{ae}t tiefster Stand: ") & Euro(Bereich("var_LiqMin").Cells(pos, 1).Value) & _
        vbLf & Txt("Die Eingaben sind gespeichert und lassen sich mit " & _
        """Variante einspielen"" wiederherstellen."), vbInformation, "Prognosemodell"
    Exit Sub
Fehler:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub

' Variante einspielen: markierte Zeile im Blatt Varianten, sonst Nummer oder Bezeichnung
Public Sub VarianteEinspielenStarten()
    Dim tabelle As Range, pos As Long, eingabe As String, bez As String
    Dim antwort As VbMsgBoxResult, n As Long
    Set tabelle = Bereich("var_Tabelle")
    If ActiveSheet.Name = tabelle.Worksheet.Name Then
        If ActiveCell.Row >= tabelle.Row And ActiveCell.Row < tabelle.Row + tabelle.Rows.Count Then
            pos = ActiveCell.Row - tabelle.Row + 1
            If CStr(tabelle.Cells(pos, 1).Value) = "" Then pos = 0
        End If
    End If
    If pos = 0 Then
        eingabe = InputBox(Txt("Nummer (1 = erste Zeile im Blatt Varianten) oder " & _
            "Bezeichnung der Variante:"), "Variante einspielen")
        If Trim$(eingabe) = "" Then Exit Sub
        If IsNumeric(eingabe) Then
            pos = CLng(eingabe)
        Else
            pos = VariantePosition(eingabe)
        End If
        If pos < 1 Or pos > tabelle.Rows.Count Then
            MsgBox Txt("Variante nicht gefunden."), vbExclamation, "Prognosemodell"
            Exit Sub
        End If
    End If
    bez = CStr(tabelle.Cells(pos, 1).Value)
    If bez = "" Or Not Gespeichert(pos) Then
        MsgBox Txt("F{ue}r diese Variante sind keine Eingaben gespeichert (festgehalten " & _
            "vor dieser Version oder aus einer {ue}bernommenen Mappe)."), vbExclamation, _
            "Prognosemodell"
        Exit Sub
    End If
    antwort = MsgBox(Txt("Variante """ & bez & """ einspielen?" & vbLf & vbLf & _
        "Parameter, Objekte, Verk{ae}ufe, Neuobjekte und Neukauf-KSt werden auf den " & _
        "festgehaltenen Stand gesetzt; Kostenstellenbl{ae}tter und Anlagen bleiben." & _
        vbLf & vbLf & "Den aktuellen Stand vorher als Variante festhalten?" & vbLf & _
        "Ja = festhalten und einspielen, Nein = ohne Festhalten einspielen"), _
        vbYesNoCancel + vbQuestion, "Variante einspielen")
    If antwort = vbCancel Then Exit Sub
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    If antwort = vbYes Then
        VarianteFesthalten "Stand vor " & bez & " (" & Format(Now, "dd.mm.yyyy hh:nn") & ")"
    End If
    n = VarianteEinspielen(pos)
    Application.ScreenUpdating = True
    MsgBox Txt("Variante """ & bez & """ eingespielt, " & n & " Felder ge{ae}ndert." & _
        vbLf & "Plausibilit{ae}tspr{ue}fung: ") & Bereich("pr_Gesamt").Value, _
        vbInformation, "Prognosemodell"
    Exit Sub
Fehler:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub
