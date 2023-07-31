Option Explicit

Const TOTAAL_VRAGEN As Integer = 296
Const CEL_RESTERENDE_TIJD As String = "C1"
Const CEL_HUIDIGE_VRAAG As String = "E1"
Const CEL_SCORE As String = "D1"
Const CEL_ANTWOORD As String = "B1"
Const WERKBLAD_SPEL As String = "3-6-9"
Const WERKBLAD_VRAGEN As String = "Vragen"
Const WERKBLAD_MENU As String = "Menu"


Dim resterendetijd As Integer
Dim NextTime As Date
Dim Vraagnummer As Integer
Dim score As Integer

Sub StartTijd_369()
    ' Stop de timer indien die al loopt
    Call StopTijd_369

    ' Reset de resterende tijd naar 20 seconden
    resterendetijd = 20
    
    ' Update de weergave van de tijd
    Call UpdateTijd_369

    ' Bepaal het moment waarop de volgende update moet plaatsvinden
    ' (over 3 seconden vanaf het huidige moment, zodat de speler enige tijd heeft om de vraag te lezen)
    NextTime = Now + TimeSerial(0, 0, 3)

    ' Stel de timer in om de "Aftellen_369" procedure aan te roepen op het moment dat is berekend
    Application.OnTime NextTime, "Aftellen_369"
End Sub
Sub UpdateTijd_369()
    ' Verlaag de resterende tijd met 1 seconde als er nog tijd over is
    If resterendetijd >= 1 Then
        resterendetijd = resterendetijd - 1
    End If

    ' Update de CEL_RESTERENDE_TIJD in het WERKBLAD_SPEL met de nieuwe resterende tijd
    Worksheets(WERKBLAD_SPEL).Range(CEL_RESTERENDE_TIJD).Value = resterendetijd
End Sub
Sub Aftellen_369()
    Dim HuidigeVraagNummer As Integer
    Dim CorrecteAntwoorden As String
    If resterendetijd > 0 Then
        Call UpdateTijd_369
        NextTime = Now + TimeValue("00:00:01")
        Application.OnTime NextTime, "Aftellen_369"
    ElseIf resterendetijd < 1 Then        ' Als de tijd op is, toon UserForm5
        With Sheets(WERKBLAD_SPEL)
            HuidigeVraagNummer = .Range(CEL_HUIDIGE_VRAAG).Value
        End With

        ' Haal het correcte antwoord op
        CorrecteAntwoorden = Sheets(WERKBLAD_VRAGEN).Cells(HuidigeVraagNummer, 2).Value
        ' Stel de tekst van het label in op het correcte antwoord
        UserForm5.Label1.Caption = "Ai, de tijd is om! Het juiste antwoord was: " & CorrecteAntwoorden
        UserForm5.Show
    End If
    If HuidigeVraagNummer = 15 Then
        Call ToonEindeQuizFormulier
        Exit Sub
    End If
    ActiveSheet.Shapes.Range(Array("Vraag")).Select
    With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
        .Transparency = 1
        .Solid
    End With

    HuidigeVraagNummer = HuidigeVraagNummer + 1
    Sheets("3-6-9").Range("E1").Value = HuidigeVraagNummer
    Sheets("3-6-9").Range("A1").Value = Sheets("Vragen").Cells(HuidigeVraagNummer, 1).Value
    Sheets("3-6-9").Range("a1").ClearContents
    resterendetijd = 20 ' Deze regel is aangepast
    Call VolgendeVraag
    Range("C13").Select
End Sub
Sub StopTijd_369()
    ' Het gebruik van On Error Resume Next zorgt ervoor dat de uitvoering van de code doorgaat, ook al treedt er een runtime-fout op
    On Error Resume Next
    
    ' Annuleer het geplande aftellen
    Application.OnTime EarliestTime:=NextTime, Procedure:="Aftellen_369", Schedule:=False
    
    ' Reset de foutafhandeling
    On Error GoTo 0
End Sub
Sub ControleerAntwoord()
    Dim HuidigeVraag As Integer
    Dim resterendetijd As Integer
    Dim antwoord As String
    Dim CorrecteAntwoorden As String
    Dim CorrectAntwoordArray() As String
    Dim i As Integer
    Dim correctGevonden As Boolean
    Dim correctIndex As Integer

    With Sheets("3-6-9")
        HuidigeVraag = .Range("E1").Value
        score = .Range("D1").Value
        antwoord = LCase(Trim(.Range("b1").Value))
    End With
    

    ' Haal de correcte antwoorden op en splits ze in een array
    CorrecteAntwoorden = Sheets("Vragen").Cells(HuidigeVraag, 2).Value
    CorrectAntwoordArray = Split(CorrecteAntwoorden, ";")
    
    ' Vergelijk het antwoord met elk correct antwoord
    correctGevonden = False
    For i = 0 To UBound(CorrectAntwoordArray)
        Dim CleanedCorrectAnswer As String
        CleanedCorrectAnswer = LCase(Trim(CorrectAntwoordArray(i)))
        
        ' Controleer of het antwoord numeriek is
        If IsNumeric(antwoord) And IsNumeric(CleanedCorrectAnswer) Then
            ' Als beide antwoorden numeriek zijn, controleer dan op een exacte match
            If antwoord = CleanedCorrectAnswer Then
                correctGevonden = True
                correctIndex = i
                Exit For
            End If
        Else
            ' Als de antwoorden niet numeriek zijn, gebruik dan de Levenshtein-functie
            If Levenshtein2(antwoord, CleanedCorrectAnswer) <= 2 Then
                correctGevonden = True
                Exit For
            End If
        End If
    Next i

    Dim UserMessage As String
    If correctGevonden Then
        UserMessage = "Goed gedaan!" & vbNewLine & "Het antwoord was inderdaad:" & vbNewLine & vbNewLine & Split(CorrecteAntwoorden, ";")(0)
        UserForm4.Label3.Caption = UserMessage
        UserForm4.Show
    If HuidigeVraag Mod 3 = 0 Then
        ' Verhoog de score
        score = score + 10
        Sheets("3-6-9").Range("d1").Value = score
    End If
    Else
        UserMessage = "Helaas!" & vbNewLine & vbNewLine & "Het juiste antwoord was: " & vbNewLine & vbNewLine & Split(CorrecteAntwoorden, ";")(0)
        UserForm4.Label3.Caption = UserMessage
        UserForm4.Show
    End If



    ' Controleer het huidige vraagnummer
    HuidigeVraag = Sheets("3-6-9").Range("E1").Value
    If HuidigeVraag <= 15 Then
    Call VolgendeVraag
    
    ' Als er al 15 vragen zijn gesteld, beëindig het spel
    ElseIf HuidigeVraag > 15 Then
    Call StopTijd_369
        Call ToonEindeQuizFormulier
        Exit Sub
        End If
End Sub
Sub VolgendeVraag()
    Dim TotaalVragen As Integer
    Dim BerichtTeksten As Variant
    Dim VraagIndex As Integer
    Dim WillekeurigeIndex As Integer
    With ActiveSheet.Shapes("Vraag").TextFrame2.TextRange.Font.Fill
        .Transparency = 1
    End With
    Range("C13").Select

    ' Initialiseer de BerichtTeksten array
    BerichtTeksten = Array("Maak je klaar voor vraag ", "Hier komt vraag ", "De volgende vraag is vraag ", "Klaar voor de volgende vraag?" & vbNewLine & "Hier komt vraag nummer ")

    ' Kies een willekeurige index
    WillekeurigeIndex = Int((UBound(BerichtTeksten) - LBound(BerichtTeksten) + 1) * Rnd + LBound(BerichtTeksten))

    TotaalVragen = Worksheets(WERKBLAD_VRAGEN).Cells(Rows.Count, 1).End(xlUp).row
    VraagIndex = Int((TotaalVragen - 1 + 1) * Rnd + 1)
    Worksheets(WERKBLAD_SPEL).Range("A1").Value = Worksheets(WERKBLAD_VRAGEN).Range("A" & Vraagnummer + 1).Value


    ' Bereid de boodschap voor de gebruiker voor
    Dim UserMessage As String
    Select Case Vraagnummer
        Case 1
            UserMessage = "We gaan beginnen!" & vbNewLine & "Hier komt vraag nummer één!"
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
        Case 3, 6, 9, 12
            UserMessage = BerichtTeksten(WillekeurigeIndex) & Vraagnummer & ". " & vbNewLine & vbNewLine & "Als je de volgende vraag goed beantwoordt, win je 10 seconden!"
            UserForm4.Label3.Caption = UserMessage
            UserForm4.Show
        Case 15
            UserMessage = "We zijn alweer toegekomen aan de laatste vraag van deze ronde!" & vbNewLine & vbNewLine & "Ook deze vraag is goed voor 10 seconden!"
            UserForm4.Label3.Caption = UserMessage
            UserForm4.Show
        Case 2, 4, 5, 7, 8, 10, 11, 13, 14
            UserMessage = BerichtTeksten(WillekeurigeIndex) & Vraagnummer & "."
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
    End Select
    ' Toon de vraag na het starten van de timer en na het sluiten van de MsgBox
    With ActiveSheet.Shapes("Vraag").TextFrame2.TextRange.Font.Fill
        .Transparency = 0
        End With
    ' Bepaal het moment waarop de volgende update moet plaatsvinden
    ' (over 3 seconden vanaf het huidige moment, zodat de speler enige tijd heeft om de vraag te lezen)
    NextTime = Now + TimeValue("00:00:03")
    
    ' Stel de timer in om de "Aftellen_369" procedure aan te roepen op het moment dat is berekend
    Application.OnTime NextTime, "Aftellen_369"
End Sub

Sub StartSpel_369()
    ' Reset de timer
    resterendetijd = 20

    ' Reset de score
    score = 60

    ' Reset het vraagnummer
    Vraagnummer = 0
    
    ' Shuffle de vragen
    Call SorteerVragenWillekeurig
    Range("a1").Select

    ' Start met de eerste vraag
    Call VolgendeVraag
End Sub

Sub ToonVragenWerkblad()
    Application.ScreenUpdating = False
    With Sheets(WERKBLAD_VRAGEN)
        .Select
    End With
    Application.ScreenUpdating = True
End Sub

Sub BeantwoordVraag_369()
    With Antwoord_369
        .Show
    End With
End Sub

Sub SorteerVragenWillekeurig()

    Dim WerkBereik As Range
    Dim WerkArray() As Variant
    Dim TempArray() As Variant
    Dim TempVariant As Variant
    Dim i As Long, j As Long
    
   ' Definieer het bereik dat gesorteerd moet worden
    Set WerkBereik = ThisWorkbook.Worksheets(WERKBLAD_VRAGEN).Range("A1:B296")

    ' Zet het bereik in een array
    WerkArray = WerkBereik.Value

    ' Maak een tijdelijke array met beide kolommen
    ReDim TempArray(1 To UBound(WerkArray, 1), 1 To UBound(WerkArray, 2))
    For i = LBound(WerkArray, 1) To UBound(WerkArray, 1)
        For j = LBound(WerkArray, 2) To UBound(WerkArray, 2)
            TempArray(i, j) = WerkArray(i, j)
        Next j
    Next i

    ' Random sorteer de tijdelijke array
    Randomize
    For i = LBound(TempArray, 1) To UBound(TempArray, 1)
        j = Int((UBound(TempArray, 1) - i + 1) * Rnd + i)
        TempVariant = TempArray(i, 1)
        TempArray(i, 1) = TempArray(j, 1)
        TempArray(j, 1) = TempVariant
        TempVariant = TempArray(i, 2)
        TempArray(i, 2) = TempArray(j, 2)
        TempArray(j, 2) = TempVariant
    Next i

    ' Zet de gesorteerde array terug in het bereik
    WerkBereik.Value = TempArray

End Sub
Sub Pas_369()
     Dim HuidigeVraagNummer As Integer
    Dim resterendetijd As Integer
    Dim antwoord As String
    Dim CorrecteAntwoorden As String
    Dim CorrecteAntwoordArray() As String
    Dim i As Integer
    Dim correctGevonden As Boolean

    With Sheets(WERKBLAD_SPEL)
        HuidigeVraagNummer = .Range(CEL_HUIDIGE_VRAAG).Value
    End With

        ' Haal het correcte antwoord op
        CorrecteAntwoorden = Sheets(WERKBLAD_VRAGEN).Cells(HuidigeVraagNummer, 2).Value
 If InStr(CorrecteAntwoorden, ";") > 0 Then
    antwoord = Split(CorrecteAntwoorden, ";")(0)
Else
    antwoord = CorrecteAntwoorden
    End If
With UserForm5
    .Label2.Caption = "Helaas!" & vbNewLine & "Het juiste antwoord was:" & vbNewLine & vbNewLine & antwoord
    .Show
End With
    ' Als HuidigeVraagNummer groter dan of gelijk aan 15 is, stop dan de quiz
    If HuidigeVraagNummer >= 15 Then
        Call ToonEindeQuizFormulier
        Exit Sub
    End If
    Call VolgendeVraag
End Sub
Sub ToonEindeQuizFormulier()
With formna369
        .StartUpPosition = 1 'Handmatig
        .Show
End With
End Sub
