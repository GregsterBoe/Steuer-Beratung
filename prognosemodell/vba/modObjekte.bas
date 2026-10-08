Attribute VB_Name = "modObjekte"
Option Explicit

' Objekte anlegen, duplizieren, entfernen; leere Prognosebloecke ausblenden.
' Jede Objektzeile hat in der Prognose einen festen Block mit Formeln. Die Makros
' schreiben nur in die Eingabezellen (obj_Eingabe) und fuegen nie Zeilen ein oder
' loeschen sie; so bleiben alle Formeln und benannten Bereiche intakt.
' Leere Eingabezellen tragen Annahmeformeln (blau); die Vorlage obj_Vorlage stellt
' sie wieder her, wenn sie geloescht wurden.

' Spalte Basisjahr der Kostenstellenblaetter (T, vorlagen.SPALTE_JAHR)
Private Const SPALTE_BASISJAHR As Long = 20

' Position der ObjektID (1 = erste Objektzeile), 0 = nicht vorhanden
Public Function ObjektPosition(ByVal objektId As String) As Long
    Dim ids As Range, i As Long
    Set ids = Bereich("obj_ID")
    For i = 1 To ids.Rows.Count
        If Trim$(CStr(ids.Cells(i, 1).Value)) = objektId Then
            ObjektPosition = i
            Exit Function
        End If
    Next i
End Function

' Erste Objektzeile, in der keine Eingabe steht; 0 = Blatt voll
Public Function FreiePosition() As Long
    Dim eingabe As Range, i As Long, j As Long, leer As Boolean
    Set eingabe = Bereich("obj_Eingabe")
    For i = 1 To eingabe.Rows.Count
        leer = True
        For j = 1 To eingabe.Columns.Count
            If CStr(eingabe.Cells(i, j).Value) <> "" Then
                leer = False
                Exit For
            End If
        Next j
        If leer Then
            FreiePosition = i
            Exit Function
        End If
    Next i
End Function

Public Function ObjektAnlegen(ByVal objektId As String) As Long
    Dim pos As Long
    objektId = Trim$(objektId)
    If objektId = "" Then Err.Raise vbObjectError + 513, "modObjekte", "ObjektID fehlt."
    If ObjektPosition(objektId) > 0 Then
        Err.Raise vbObjectError + 514, "modObjekte", "ObjektID " & objektId & " gibt es schon."
    End If
    pos = FreiePosition()
    If pos = 0 Then
        Err.Raise vbObjectError + 515, "modObjekte", "Kein freier Platz im Blatt Objekte."
    End If
    Bereich("obj_ID").Cells(pos, 1).Value = objektId
    ObjektAnlegen = pos
End Function

' Alle Eingaben eines Objekts in eine freie Zeile kopieren, mit neuer ObjektID
Public Function ObjektDuplizieren(ByVal quelleId As String, ByVal neueId As String) As Long
    Dim quelle As Long, ziel As Long, eingabe As Range
    quelle = ObjektPosition(Trim$(quelleId))
    If quelle = 0 Then
        Err.Raise vbObjectError + 516, "modObjekte", "ObjektID " & quelleId & " nicht gefunden."
    End If
    ziel = ObjektAnlegen(neueId)
    Set eingabe = Bereich("obj_Eingabe")
    ' FormulaR1C1: Annahmeformeln bleiben Formeln und beziehen sich auf die neue Zeile
    eingabe.Rows(ziel).FormulaR1C1 = eingabe.Rows(quelle).FormulaR1C1
    eingabe.Cells(ziel, 1).Value = Trim$(neueId)
    ObjektDuplizieren = ziel
End Function

' Eingaben eines Objekts leeren; Formeln und Prognoseblock bleiben stehen
Public Sub ObjektEntfernen(ByVal objektId As String)
    Dim pos As Long
    pos = ObjektPosition(Trim$(objektId))
    If pos = 0 Then
        Err.Raise vbObjectError + 516, "modObjekte", "ObjektID " & objektId & " nicht gefunden."
    End If
    Bereich("obj_Eingabe").Rows(pos).ClearContents
    AnnahmenWiederherstellen pos
End Sub

' Leere Eingabezellen mit Annahmeformel aus der Vorlagezeile fuellen; pos = 0: alle
' Objektzeilen. Liefert die Anzahl wiederhergestellter Zellen.
Public Function AnnahmenWiederherstellen(Optional ByVal pos As Long = 0) As Long
    Dim eingabe As Range, vorlage As Range, i As Long, j As Long, von As Long, bis As Long
    Set eingabe = Bereich("obj_Eingabe")
    Set vorlage = Bereich("obj_Vorlage")
    von = 1: bis = eingabe.Rows.Count
    If pos > 0 Then von = pos: bis = pos
    For i = von To bis
        For j = 1 To vorlage.Columns.Count
            If vorlage.Cells(1, j).HasFormula And Not eingabe.Cells(i, j).HasFormula Then
                If CStr(eingabe.Cells(i, j).Value) = "" Then
                    eingabe.Cells(i, j).FormulaR1C1 = vorlage.Cells(1, j).FormulaR1C1
                    AnnahmenWiederherstellen = AnnahmenWiederherstellen + 1
                End If
            End If
        Next j
    Next i
End Function

' Kostenstellenblatt fuehrt: Basiswerte, die im Blatt Objekte ueber die Verknuepfung
' getippt wurden, ins Kostenstellenblatt (Spalte Basisjahr) schreiben und die
' Verknuepfung wiederherstellen. Weitere Ausgaben: die Differenz geht auf 1260 Sonstige
' Kosten. Geleerte Felder bekommen nur die Verknuepfung zurueck. Liefert die Anzahl Felder.
Public Function InKostenstelleUebernehmen() As Long
    Dim felder As Variant, f As Variant, eingabe As Range, verkn As Range, blaetter As Range
    Dim i As Long, ziel As Range, neu As Double
    felder = Array(Array("obj_MieteBasis", "obk_miete", 1020), _
                   Array("obj_ErhBasis", "obk_erhaltung", 1250), _
                   Array("obj_EinnBasis", "obk_weitere_einnahmen", 1090), _
                   Array("obj_AusgBasis", "obk_weitere_ausgaben", 1260), _
                   Array("obj_AfABWA", "obk_afa_bwa", 1240), _
                   Array("obj_ZinsBasis", "obk_zinsen", 1310))
    Set blaetter = Bereich("obj_KStBlatt")
    For Each f In felder
        Set eingabe = Bereich(f(0))
        Set verkn = Bereich(f(1))
        For i = 1 To eingabe.Rows.Count
            If verkn.Cells(i, 1).HasFormula And Not eingabe.Cells(i, 1).HasFormula _
                    And CStr(blaetter.Cells(i, 1).Value) <> "" Then
                If CStr(eingabe.Cells(i, 1).Value) <> "" Then
                    Set ziel = BasisjahrZelle(CStr(blaetter.Cells(i, 1).Value), CLng(f(2)))
                    neu = CDbl(eingabe.Cells(i, 1).Value)
                    ' Differenz, damit bei weiteren Ausgaben die Kostenarten bleiben
                    ziel.Value = Zahl(ziel.Value) + neu - Zahl(verkn.Cells(i, 1).Value)
                End If
                eingabe.Cells(i, 1).Formula = verkn.Cells(i, 1).Formula
                InKostenstelleUebernehmen = InKostenstelleUebernehmen + 1
            End If
        Next i
    Next f
    Application.Calculate
End Function

Public Sub InKostenstelleUebernehmenStarten()
    Dim n As Long
    On Error GoTo Fehler
    n = InKostenstelleUebernehmen()
    MsgBox n & Txt(" Werte aus dem Blatt Objekte in die Kostenstellenbl{ae}tter " & _
        "{ue}bernommen, Verkn{ue}pfungen wiederhergestellt."), vbInformation, "Prognosemodell"
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub

' Zelle der BWA-Nr. nr in der Spalte Basisjahr eines Kostenstellenblatts
Private Function BasisjahrZelle(ByVal blatt As String, ByVal nr As Long) As Range
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(blatt)
    For z = 1 To 200
        If IsNumeric(ws.Cells(z, 2).Value) And CStr(ws.Cells(z, 2).Value) <> "" Then
            If CLng(ws.Cells(z, 2).Value) = nr Then
                Set BasisjahrZelle = ws.Cells(z, SPALTE_BASISJAHR)
                Exit Function
            End If
        End If
    Next z
    Err.Raise vbObjectError + 517, "modObjekte", "BWA-Nr. " & nr & " fehlt im Blatt " & blatt
End Function

Private Function Zahl(ByVal wert As Variant) As Double
    If IsNumeric(wert) And CStr(wert) <> "" Then Zahl = CDbl(wert)
End Function

Public Sub AnnahmenWiederherstellenStarten()
    Dim n As Long
    n = AnnahmenWiederherstellen()
    MsgBox n & Txt(" leere Felder wieder mit Annahmen gef{ue}llt (blau)."), vbInformation, _
        "Prognosemodell"
End Sub

' Prognosezeilen ohne ObjektID (leere Objekt- und Neuobjektzeilen) aus- oder einblenden
Public Sub LeereBloeckeAusblenden(ByVal ausblenden As Boolean)
    Dim ids As Range, ws As Worksheet, i As Long, start As Long, leer As Boolean
    Set ids = Bereich("prg_ID")
    Set ws = ids.Worksheet
    ids.EntireRow.Hidden = False
    If ausblenden Then
        For i = 1 To ids.Rows.Count + 1
            leer = False
            If i <= ids.Rows.Count Then leer = (CStr(ids.Cells(i, 1).Value) = "")
            If leer And start = 0 Then start = i
            If Not leer And start > 0 Then
                ws.Range(ids.Cells(start, 1), ids.Cells(i - 1, 1)).EntireRow.Hidden = True
                start = 0
            End If
        Next i
    End If
End Sub

Public Function LeereBloeckeAusgeblendet() As Boolean
    Dim ids As Range, i As Long
    Set ids = Bereich("prg_ID")
    For i = 1 To ids.Rows.Count
        If ids.Cells(i, 1).EntireRow.Hidden Then
            LeereBloeckeAusgeblendet = True
            Exit Function
        End If
    Next i
End Function

' ---- Bedienung ueber Schaltflaechen ----

Public Sub ObjektAnlegenStarten()
    Dim objektId As String
    objektId = InputBox(Txt("ObjektID f{ue}r das neue Objekt:"), "Objekt anlegen", VorschlagId())
    If objektId = "" Then Exit Sub
    On Error GoTo Fehler
    ZeileZeigen ObjektAnlegen(objektId)
    MsgBox Txt("Objekt angelegt. Bitte die gelben Pflichtfelder ausf{ue}llen; die Spalte " & _
        "Status zeigt, was noch fehlt."), vbInformation, "Prognosemodell"
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub

Public Sub ObjektDuplizierenStarten()
    Dim quelleId As String, neueId As String
    quelleId = AktiveObjektId()
    If quelleId = "" Then quelleId = InputBox("Welches Objekt duplizieren? ObjektID:", _
        "Objekt duplizieren")
    If quelleId = "" Then Exit Sub
    neueId = InputBox(Txt("Neue ObjektID f{ue}r die Kopie von ") & quelleId & ":", _
        "Objekt duplizieren", VorschlagId())
    If neueId = "" Then Exit Sub
    On Error GoTo Fehler
    ZeileZeigen ObjektDuplizieren(quelleId, neueId)
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub

Public Sub ObjektEntfernenStarten()
    Dim objektId As String, frage As String
    objektId = AktiveObjektId()
    If objektId = "" Then objektId = InputBox("Welches Objekt entfernen? ObjektID:", _
        "Objekt entfernen")
    If objektId = "" Then Exit Sub
    frage = Txt("Eingaben von Objekt ") & objektId & Txt(" l{oe}schen?")
    If Application.WorksheetFunction.CountIf(Bereich("vk_ID"), objektId) > 0 Then
        frage = frage & vbLf & Txt("Ein Verkauf nennt das Objekt noch; er bleibt stehen und " & _
            "wird als Fehler gemeldet.")
    End If
    If MsgBox(frage, vbYesNo + vbQuestion, "Objekt entfernen") <> vbYes Then Exit Sub
    On Error GoTo Fehler
    ObjektEntfernen objektId
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "Prognosemodell"
End Sub

Public Sub LeereBloeckeUmschaltenStarten()
    Application.ScreenUpdating = False
    LeereBloeckeAusblenden Not LeereBloeckeAusgeblendet()
    Application.ScreenUpdating = True
    Bereich("prg_ID").Worksheet.Activate
End Sub

' ObjektID der Zeile, in der die aktive Zelle auf dem Objektblatt steht
Private Function AktiveObjektId() As String
    Dim ids As Range, pos As Long
    Set ids = Bereich("obj_ID")
    If ActiveSheet.Name <> ids.Worksheet.Name Then Exit Function
    pos = ActiveCell.Row - ids.Row + 1
    If pos >= 1 And pos <= ids.Rows.Count Then AktiveObjektId = Trim$(CStr(ids.Cells(pos, 1).Value))
End Function

Private Function VorschlagId() As String
    Dim n As Long
    n = Application.WorksheetFunction.CountA(Bereich("obj_ID")) + 1
    Do While ObjektPosition("OBJ-" & Format(n, "000")) > 0
        n = n + 1
    Loop
    VorschlagId = "OBJ-" & Format(n, "000")
End Function

Private Sub ZeileZeigen(ByVal pos As Long)
    Dim ids As Range
    Set ids = Bereich("obj_ID")
    ids.Worksheet.Activate
    ids.Cells(pos, 2).Select
End Sub
