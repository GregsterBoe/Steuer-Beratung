Attribute VB_Name = "modOptimierung"
Option Explicit

' Optimierung (Blatt Optimierung): Das Makro rechnet nicht selbst. Es setzt je Neuobjekt
' Kaufjahr, "Paragraph 6b nur bis QNG-Grenze" und Kreditquote, loest die Neuberechnung aus
' und liest Ziel und Nebenbedingungen aus den Formeln des Blatts Optimierung:
'   Ziel            opt_Ziel (Endvermoegen nach latenter Steuer, Plan A)
'   zulaessig, wenn opt_LiqJahre = 0, opt_KDDJahre = 0 und keine neuen Fehler (pr_Fehler)
' Suche: koordinatenweise, je Runde jeder Hebel jedes Neuobjekts einmal; ein Kandidat wird
' nur behalten, wenn er besser ist. Unzulaessige Loesungen werden nach Abstand zur
' Zulaessigkeit verglichen, damit die Suche auch aus einem unzulaessigen Start herausfindet.

Private Const JA As String = "ja"
Private Const NEIN As String = "nein"
Private Const KREDIT As String = "Kredit"
Private Const UNZULAESSIG As Double = -1E+15
Private Const BESSER_AB As Double = 1   ' Euro: kleinere Verbesserungen zaehlen nicht

Private mLaeufe As Long
Private mFehlerStart As Double
Private mZulaessig As Boolean
Private mBestWert As Double

' ---- Hilfen ----

Private Function Zahl(ByVal name As String) As Double
    Dim v As Variant
    v = Bereich(name).Cells(1, 1).Value
    If IsNumeric(v) And CStr(v) <> "" Then Zahl = CDbl(v)
End Function

Private Function Klein(ByVal name As String) As String
    Klein = LCase$(Trim$(CStr(Bereich(name).Cells(1, 1).Value)))
End Function

Private Function Zelle(ByVal name As String, ByVal i As Long) As Range
    Set Zelle = Bereich(name).Cells(i, 1)
End Function

Private Function Inhalt(ByVal z As Range) As String
    Inhalt = CStr(z.Formula)
End Function

' Eingabe zuruecksetzen (Formel oder Konstante, leer = leeren)
Private Sub Setzen(ByVal z As Range, ByVal formel As String)
    If formel = "" Then
        z.ClearContents
    Else
        z.Formula = formel
    End If
End Sub

Private Sub Status(ByVal s As String)
    On Error Resume Next
    Application.StatusBar = s
    On Error GoTo 0
End Sub

' Neu rechnen und bewerten: zulaessig = Ziel, sonst stark negativ nach Abstand
Private Function Bewertung() As Double
    Dim fehler As Double, kdd As Double, liq As Double
    NeuBerechnen
    mLaeufe = mLaeufe + 1
    fehler = Zahl("pr_Fehler") - mFehlerStart
    If fehler < 0 Then fehler = 0
    kdd = Zahl("opt_KDDJahre")
    liq = Zahl("opt_LiqJahre")
    mZulaessig = (fehler = 0 And kdd = 0 And liq = 0)
    If mZulaessig Then
        Bewertung = Zahl("opt_Ziel")
    Else
        Bewertung = UNZULAESSIG - fehler * 1E+13 - kdd * 1E+10 - Zahl("opt_Fehlbetrag")
    End If
End Function

' Nebenbedingungen ohne Pruefung auf Fehler (fuer die tragbare Ausschuettung)
Private Function NebenbedingungenErfuellt() As Boolean
    NeuBerechnen
    mLaeufe = mLaeufe + 1
    NebenbedingungenErfuellt = (Zahl("opt_LiqJahre") = 0 And Zahl("opt_KDDJahre") = 0)
End Function

' ---- Hebel je Neuobjekt ----
' Ein Hebel ist eine Liste von Kandidaten; Kandidat 0 ist der Ausgang (Formeln wie erfasst).

Private Function Kandidaten(ByVal art As String, ByVal i As Long, ByRef hinweis As String) As Variant
    Dim liste() As Variant, n As Long, q As Double, j As Long, s As Long, kj As Long
    Dim schritt As Double, bis As Double
    ReDim liste(0 To 0)
    liste(0) = Empty
    Select Case art
        Case "Jahr"
            If Klein("opt_HebelJahr") <> JA Then GoTo Ende
            If Not IsNumeric(Zelle("ne_Kaufjahr", i).Value) Then GoTo Ende
            kj = CLng(Zelle("ne_Kaufjahr", i).Value)
            s = CLng(Zahl("opt_JahrSpielraum"))
            For j = kj - s To kj + s
                If j <> kj And j >= Zahl("par_Startjahr") And j <= Zahl("par_Endjahr") Then
                    n = n + 1
                    ReDim Preserve liste(0 To n)
                    liste(n) = j
                End If
            Next j
        Case "QNG"
            If Klein("opt_Hebel6b") <> JA Then GoTo Ende
            If LCase$(Trim$(CStr(Zelle("ne_QNG", i).Value))) <> JA Then GoTo Ende
            If Trim$(CStr(Zelle("ne_Quelle", i).Value)) = "" And _
                    Trim$(CStr(Zelle("ne_Quelle2", i).Value)) = "" And _
                    Trim$(CStr(Zelle("ne_Quelle3", i).Value)) = "" Then GoTo Ende
            ReDim Preserve liste(0 To 1)
            If LCase$(Trim$(CStr(Zelle("ne_6bQNG", i).Value))) = JA Then
                liste(1) = NEIN
            Else
                liste(1) = JA
            End If
        Case "Quote"
            If Klein("opt_HebelKredit") <> JA Then GoTo Ende
            If Trim$(CStr(Zelle("ne_KreditZins", i).Value)) = "" Then
                hinweis = hinweis & Txt("Kreditquote: Zinssatz fehlt im Blatt Neuobjekte. ")
                GoTo Ende
            End If
            schritt = Zahl("opt_KreditSchritt")
            bis = Zahl("opt_KreditMax")
            If schritt <= 0 Then schritt = 0.1
            q = Zahl("opt_KreditMin")
            Do While q <= bis + 0.000001
                n = n + 1
                ReDim Preserve liste(0 To n)
                liste(n) = Round(q, 6)
                q = q + schritt
            Loop
    End Select
Ende:
    Kandidaten = liste
End Function

' Ausgang eines Hebels: Formeln der beteiligten Zellen
Private Function Ausgang(ByVal art As String, ByVal i As Long) As Variant
    Select Case art
        Case "Jahr": Ausgang = Array(Inhalt(Zelle("ne_Kaufjahr", i)))
        Case "QNG": Ausgang = Array(Inhalt(Zelle("ne_6bQNG", i)))
        Case "Quote": Ausgang = Array(Inhalt(Zelle("ne_FinArt", i)), _
            Inhalt(Zelle("ne_KreditBetrag", i)), Inhalt(Zelle("ne_KreditQuote", i)))
    End Select
End Function

' Kandidat k setzen; k = 0 stellt den Ausgang wieder her
Private Sub Anwenden(ByVal art As String, ByVal i As Long, ByVal liste As Variant, _
        ByVal ausgangswerte As Variant, ByVal k As Long)
    If k = 0 Then
        Select Case art
            Case "Jahr": Setzen Zelle("ne_Kaufjahr", i), CStr(ausgangswerte(0))
            Case "QNG": Setzen Zelle("ne_6bQNG", i), CStr(ausgangswerte(0))
            Case "Quote"
                Setzen Zelle("ne_FinArt", i), CStr(ausgangswerte(0))
                Setzen Zelle("ne_KreditBetrag", i), CStr(ausgangswerte(1))
                Setzen Zelle("ne_KreditQuote", i), CStr(ausgangswerte(2))
        End Select
        Exit Sub
    End If
    Select Case art
        Case "Jahr": Zelle("ne_Kaufjahr", i).Value = liste(k)
        Case "QNG": Zelle("ne_6bQNG", i).Value = liste(k)
        Case "Quote"
            Zelle("ne_FinArt", i).Value = KREDIT
            Zelle("ne_KreditBetrag", i).ClearContents
            Zelle("ne_KreditQuote", i).Value = liste(k)
    End Select
End Sub

' Einen Hebel durchprobieren, ausgehend vom bisher besten Kandidaten aktuell; liefert den
' besten Kandidaten. mBestWert haelt die beste Bewertung.
Private Function HebelPruefen(ByVal art As String, ByVal i As Long, ByVal aktuell As Long, _
        ByVal liste As Variant, ByVal ausgangswerte As Variant) As Long
    Dim k As Long, wert As Double, best As Long
    best = aktuell
    For k = 0 To UBound(liste)
        If k <> aktuell Then
            Anwenden art, i, liste, ausgangswerte, k
            wert = Bewertung()
            If wert > mBestWert + BESSER_AB Then
                mBestWert = wert
                best = k
            End If
        End If
    Next k
    Anwenden art, i, liste, ausgangswerte, best
    HebelPruefen = best
End Function

' ---- Optimieren ----

' Sucht die beste zulaessige Kombination und laesst sie in den Eingaben stehen.
' Liefert die Anzahl Hebel, die vom Ausgang abweichen.
Public Function Optimieren() As Long
    Dim ids As Range, i As Long, n As Long, a As Long, runde As Long, runden As Long
    Dim arten As Variant, art As String, hinweis As String
    Dim neuIdx() As Long, kand() As Variant, ausg() As Variant, bestK() As Long
    Dim verbessert As Boolean, aenderungen As Long, neu As Long
    Dim vorherZiel As Variant, vorherLiq As Variant

    arten = Array("Jahr", "QNG", "Quote")
    mLaeufe = 0
    NeuBerechnen
    mFehlerStart = Zahl("pr_Fehler")
    vorherZiel = Bereich("opt_Ziel").Value
    vorherLiq = Bereich("opt_LiqMin").Value
    If Klein("opt_Festhalten") = JA Then Festhalten "Optimierung Ausgang"

    ' Neuobjekte, die rechnen und optimiert werden sollen
    Set ids = Bereich("ne_ID")
    Bereich("opt_VorQuote").ClearContents
    Bereich("opt_VorJahr").ClearContents
    Bereich("opt_Vor6b").ClearContents
    Bereich("opt_Hinweis").ClearContents
    For i = 1 To ids.Rows.Count
        If Trim$(CStr(ids.Cells(i, 1).Value)) <> "" Then
            Zelle("opt_VorQuote", i).Value = Bereich("opt_Quote").Cells(i, 1).Value
            Zelle("opt_VorJahr", i).Value = Zelle("ne_Kaufjahr", i).Value
            Zelle("opt_Vor6b", i).Value = Bereich("opt_QNGGrenze").Cells(i, 1).Value
            If LCase$(Trim$(CStr(Zelle("opt_Mit", i).Value))) = NEIN Then
                Zelle("opt_Hinweis", i).Value = Txt("nicht optimiert (optimieren = nein)")
            ElseIf CStr(Zelle("ne_Gueltig", i).Value) <> "1" Then
                Zelle("opt_Hinweis", i).Value = Txt("rechnet nicht mit (Status bzw. Kauf aktiv)")
            Else
                n = n + 1
                ReDim Preserve neuIdx(1 To n)
                neuIdx(n) = i
            End If
        End If
    Next i

    mBestWert = Bewertung()
    If n > 0 Then
        ReDim kand(1 To n, 0 To 2)
        ReDim ausg(1 To n, 0 To 2)
        ReDim bestK(1 To n, 0 To 2)
        For a = 1 To n
            hinweis = ""
            For i = 0 To 2
                kand(a, i) = Kandidaten(CStr(arten(i)), neuIdx(a), hinweis)
                ausg(a, i) = Ausgang(CStr(arten(i)), neuIdx(a))
            Next i
            If hinweis <> "" Then Zelle("opt_Hinweis", neuIdx(a)).Value = Trim$(hinweis)
        Next a
        runden = CLng(Zahl("opt_Runden"))
        If runden < 1 Then runden = 1
        For runde = 1 To runden
            verbessert = False
            For a = 1 To n
                For i = 0 To 2
                    If UBound(kand(a, i)) > 0 Then
                        art = CStr(arten(i))
                        Status Txt("Optimierung: Runde " & runde & ", " & _
                            ids.Cells(neuIdx(a), 1).Value & ", " & art & ", Rechenlauf " & mLaeufe)
                        neu = HebelPruefen(art, neuIdx(a), bestK(a, i), kand(a, i), ausg(a, i))
                        If neu <> bestK(a, i) Then
                            bestK(a, i) = neu
                            verbessert = True
                        End If
                    End If
                Next i
            Next a
            If Not verbessert Then Exit For
        Next runde
        For a = 1 To n
            For i = 0 To 2
                If bestK(a, i) <> 0 Then aenderungen = aenderungen + 1
            Next i
        Next a
    End If
    Bewertung
    Status ""

    Bereich("opt_ErgZeit").Value = Now
    Bereich("opt_ErgVorher").Value = vorherZiel
    Bereich("opt_ErgNachher").Value = Bereich("opt_Ziel").Value
    Bereich("opt_ErgLiqVorher").Value = vorherLiq
    Bereich("opt_ErgLiqNachher").Value = Bereich("opt_LiqMin").Value
    Bereich("opt_ErgErfuellt").Value = IIf(mZulaessig, JA, NEIN)
    Bereich("opt_ErgAenderungen").Value = aenderungen
    Bereich("opt_ErgLaeufe").Value = mLaeufe
    If n = 0 Then
        Bereich("opt_ErgMeldung").Value = Txt("Kein Neuobjekt zum Optimieren.")
    ElseIf Not mZulaessig Then
        Bereich("opt_ErgMeldung").Value = Txt("Nebenbedingungen mit diesen Hebeln nicht " & _
            "erf{ue}llbar; beste N{ae}herung eingetragen. Aussch{ue}ttung senken (Makro " & _
            "Tragbare Aussch{ue}ttung), Spielraum erweitern oder weitere K{ae}ufe erfassen.")
    ElseIf aenderungen = 0 Then
        Bereich("opt_ErgMeldung").Value = Txt("Eingaben sind bereits die beste gefundene " & _
            "Kombination.")
    Else
        Bereich("opt_ErgMeldung").Value = Txt("Beste zul{ae}ssige Kombination eingetragen.")
    End If
    If Klein("opt_Festhalten") = JA Then Festhalten "Optimierung Ergebnis"
    Optimieren = aenderungen
End Function

Private Sub Festhalten(ByVal bezeichnung As String)
    On Error Resume Next
    VarianteFesthalten bezeichnung
    On Error GoTo 0
End Sub

' ---- Tragbare Ausschuettung ----

' Hoechste Ausschuettung p. a. (auf 1.000 Euro abgerundet), bei der Mindestliquiditaet und
' Mindest-Kapitaldienstdeckung mit den aktuellen Eingaben halten. Ab-Jahr und Steigerung
' wie Parameterblatt; der Parameter selbst bleibt unveraendert.
Public Function TragbareAusschuettung() As Double
    Dim z As Range, vorherFormel As String, lo As Double, hi As Double, mitte As Double
    Set z = Bereich("par_Ausschuettung")
    vorherFormel = Inhalt(z)
    mLaeufe = 0
    z.Value = 0
    If Not NebenbedingungenErfuellt() Then
        lo = 0
    Else
        lo = 0
        hi = Zahl("opt_Ausschuettung")
        If hi < 100000 Then hi = 100000
        z.Value = hi
        Do While NebenbedingungenErfuellt()
            lo = hi
            hi = hi * 2
            If hi > 1000000000# Then Exit Do
            z.Value = hi
        Loop
        Do While hi - lo > 1000
            mitte = (lo + hi) / 2
            z.Value = mitte
            If NebenbedingungenErfuellt() Then lo = mitte Else hi = mitte
        Loop
    End If
    Setzen z, vorherFormel
    NeuBerechnen
    TragbareAusschuettung = Int(lo / 1000) * 1000
    Bereich("opt_Tragbar").Value = TragbareAusschuettung
    Bereich("opt_TragbarZeit").Value = Now
End Function

' ---- Bedienung ueber Schaltflaechen ----

Public Sub OptimierenStarten()
    Dim antwort As VbMsgBoxResult, n As Long
    antwort = MsgBox(Txt("Kaufjahr, Paragraph 6b nur bis QNG-Grenze und Kreditquote der " & _
        "Neuobjekte optimieren?" & vbLf & vbLf & "Das Makro {ae}ndert die Eingaben im Blatt " & _
        "Neuobjekte und rechnet jede Kombination durch; das kann einige Minuten dauern. " & _
        "Mit ""Ausgang und Ergebnis festhalten = ja"" l{ae}sst sich der Ausgang im Blatt " & _
        "Varianten wieder einspielen."), vbOKCancel + vbQuestion, "Optimierung")
    If antwort <> vbOK Then Exit Sub
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    n = Optimieren()
    Application.ScreenUpdating = True
    Bereich("opt_Ziel").Worksheet.Activate
    MsgBox CStr(Bereich("opt_ErgMeldung").Value) & vbLf & vbLf & _
        Txt("Endverm{oe}gen A vorher: ") & Euro(Bereich("opt_ErgVorher").Value) & vbLf & _
        Txt("Endverm{oe}gen A nachher: ") & Euro(Bereich("opt_ErgNachher").Value) & vbLf & _
        Txt("ge{ae}nderte Hebel: ") & n & ", " & Txt("Rechenl{ae}ufe: ") & mLaeufe, _
        vbInformation, "Optimierung"
    Exit Sub
Fehler:
    Application.ScreenUpdating = True
    Status ""
    MsgBox Err.Description, vbExclamation, "Optimierung"
End Sub

Public Sub TragbareAusschuettungStarten()
    Dim betrag As Double
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    betrag = TragbareAusschuettung()
    Application.ScreenUpdating = True
    MsgBox Txt("Tragbare Aussch{ue}ttung p. a. mit den aktuellen Eingaben: ") & Euro(betrag) & _
        vbLf & Txt("(Mindestliquidit{ae}t und Mindest-Kapitaldienstdeckung laut " & _
        "Parameterblatt; der Parameter Aussch{ue}ttung bleibt unver{ae}ndert.)"), _
        vbInformation, "Optimierung"
    Exit Sub
Fehler:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Optimierung"
End Sub
