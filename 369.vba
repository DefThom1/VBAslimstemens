Option Explicit

Const TOTAAL_VRAGEN As Integer = 349
Const CEL_RESTERENDE_TIJD As String = "C1"
Const CEL_HUIDIGE_VRAAG As String = "E1"
Const CEL_SCORE As String = "D1"
Const CEL_ANTWOORD As String = "B1"
Const WERKBLAD_SPEL As String = "3-6-9"
Const WERKBLAD_VRAGEN As String = "Vragen"
Const WERKBLAD_MENU As String = "Menu"


Dim resterendetijd As Integer
Dim VolgendeTijd As Date
Dim Vraagnummer As Integer
Dim score As Integer
Dim IsAftellenActief As Boolean

Sub StartTijd_369()
    ' Stop de timer indien die al loopt
    Call StopTijd_369

    ' Reset de resterende tijd naar 30 seconden
    resterendetijd = 30
    Call UpdateTijd_369
        VolgendeTijd = Now + TimeSerial(0, 0, 1)
    Application.OnTime VolgendeTijd, "Aftellen_369"
End Sub
Sub UpdateTijd_369()
    ' Verlaag de resterende tijd met 1 seconde als er nog tijd over is
    If resterendetijd > 0 Then
        resterendetijd = resterendetijd - 1
    End If

    ' Update de CEL_RESTERENDE_TIJD in het WERKBLAD_SPEL met de nieuwe resterende tijd
    Worksheets(WERKBLAD_SPEL).Range(CEL_RESTERENDE_TIJD).Value = resterendetijd
End Sub

Sub Aftellen_369()
    ' Controleer of de aftelling al actief is
    If IsAftellenActief Then Exit Sub

    ' Zet de aftelling op actief
    IsAftellenActief = True
    Dim HuidigeVraagnummer As Integer
    Dim CorrecteAntwoorden As String
    If resterendetijd > 0 Then
    Call UpdateTijd_369
        VolgendeTijd = Now + TimeSerial(0, 0, 1)
    Application.OnTime VolgendeTijd, "Aftellen_369"
  If resterendetijd <= 1 Then
        ' Haal het huidige vraagnummer
        With Sheets(WERKBLAD_SPEL)
            HuidigeVraagnummer = .Range(CEL_HUIDIGE_VRAAG).Value
        End With

        ' Haal het correcte antwoord op
        CorrecteAntwoorden = Sheets(WERKBLAD_VRAGEN).Cells(HuidigeVraagnummer, 2).Value
End If
    ' Als er al 15 vragen zijn gesteld, beëindig het spel
    If HuidigeVraagnummer >= 15 Then
        formna369.Show
        Exit Sub
ActiveSheet.Shapes.Range(Array("Vraag369")).Select
    With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
        .Transparency = 1
        .Solid
        
        HuidigeVraagnummer = HuidigeVraagnummer + 1
Sheets("3-6-9").Range("E1").Value = HuidigeVraagnummer
Sheets("3-6-9").Range("A1").Value = Sheets("Vragen").Cells(HuidigeVraagnummer, 1).Value
Sheets("3-6-9").Range("a1").ClearContents
Sheets("3-6-9").Range("B1").ClearContents
Sheets("3-6-9").Range("F1").ClearContents
resterendetijd = 30 ' Deze regel is aangepast
Range("C13").Select
StopTijd_369
Call VolgendeVraag
End With
Range("C13").Select
End If
End If
    ' Zet de aftelling op inactief aan het einde van de subroutine
    IsAftellenActief = False
End Sub

Sub Tijd_369()
    Dim HuidigeVraagnummer As Integer
        Dim CorrecteAntwoorden As String
        
    HuidigeVraagnummer = Sheets("3-6-9").Range("E1").Value
 ' Haal het correcte antwoord op
        CorrecteAntwoorden = Sheets(WERKBLAD_VRAGEN).Cells(HuidigeVraagnummer, 2).Value
        
Tijdvoorbij369.Label1.Caption = "Helaas, de tijd is om!" & vbNewLine & "Het antwoord was:" & vbNewLine & vbNewLine & Split(CorrecteAntwoorden, ";")(0)
       
        ' Anders, ga naar de volgende vraag
 HuidigeVraagnummer = HuidigeVraagnummer + 1
        With Sheets("3-6-9")
            .Range("E1").Value = HuidigeVraagnummer
            .Range("A1").Value = Sheets("Vragen").Cells(HuidigeVraagnummer, 1).Value
            .Range("A1").ClearContents
            .Range("B1").ClearContents
            .Range("C1").Value = 30
        End With
    Call VolgendeVraag
End Sub



Sub StopTijd_369()
    ' Het gebruik van On Error Resume Next zorgt ervoor dat de uitvoering van de code doorgaat, ook al treedt er een runtime-fout op
    On Error Resume Next
    
    ' Annuleer het geplande aftellen
    Application.OnTime EarliestTime:=VolgendeTijd, Procedure:="Aftellen_369", Schedule:=False
    
    ' Reset de foutafhandeling
    On Error GoTo 0
End Sub
Sub ControleerAntwoord()
    Dim HuidigeVraagnummer As Integer
    Dim resterendetijd As Integer
    Dim antwoord As String
    Dim CorrecteAntwoorden As String
    Dim CorrectAntwoordArray() As String
    Dim i As Integer
    Dim correctGevonden As Boolean
    Dim correctIndex As Integer

    With Sheets("3-6-9")
        HuidigeVraagnummer = .Range("E1").Value
        score = .Range("D1").Value
        antwoord = LCase(Trim(.Range("b1").Value))
    End With
    ' Stop de timer
    Call StopTijd_369
    CorrecteAntwoorden = Sheets("Vragen").Cells(HuidigeVraagnummer, 2).Value
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
        UserMessage = "Goed gedaan!" & vbNewLine & vbNewLine & "Het antwoord was inderdaad:" & vbNewLine & Trim(CorrectAntwoordArray(correctIndex))
        Juistantwoord369.Label1.Caption = UserMessage
        Juistantwoord369.Show
        If HuidigeVraagnummer Mod 3 = 0 Then
        ' Verhoog de score
        score = score + 10
        Sheets("3-6-9").Range("d1").Value = score
    End If
    Else
        UserMessage = "Helaas!" & vbNewLine & vbNewLine & "Het antwoord had moeten zijn:" & vbNewLine & Trim(CorrectAntwoordArray(correctIndex))
        Onjuistantwoord369.Label1.Caption = UserMessage
        Onjuistantwoord369.Show
    End If
    ' Controleer het huidige vraagnummer
    HuidigeVraagnummer = Sheets("3-6-9").Range("E1").Value

    ' Als er al 15 vragen zijn gesteld, beëindig het spel
    If HuidigeVraagnummer >= 15 Then
    Call StopTijd_369
formna369.Show
        Exit Sub
    Else
        ' Anders, ga naar de volgende vraag
        HuidigeVraagnummer = HuidigeVraagnummer + 1
        With Sheets("3-6-9")
            .Range("E1").Value = HuidigeVraagnummer
            .Range("A1").Value = Sheets("Vragen").Cells(HuidigeVraagnummer, 1).Value
            .Range("A1").ClearContents
            .Range("B1").ClearContents
            .Range("C1").Value = 30
        End With
    End If
    Call VolgendeVraag
End Sub
Sub VolgendeVraag()
    Dim TotaalVragen As Integer
    Dim BerichtTeksten As Variant
    Dim VraagIndex As Integer
    Dim WillekeurigeIndex As Integer
    With ActiveSheet.Shapes("Vraag369").TextFrame2.TextRange.Font.Fill
        .Transparency = 1
    End With
    Range("C13").Select
        Call StopTijd_369
    ' Initialiseer de BerichtTeksten array
    BerichtTeksten = Array("Maak je klaar voor vraag ", "Hier komt vraag ", "De volgende vraag is vraag ", "Klaar voor de volgende vraag?" & vbNewLine & "Hier komt vraag nummer ")

' Kies een willekeurige index
WillekeurigeIndex = Int((UBound(BerichtTeksten) - LBound(BerichtTeksten) + 1) * Rnd + LBound(BerichtTeksten))

    TotaalVragen = Worksheets(WERKBLAD_VRAGEN).Cells(Rows.Count, 1).End(xlUp).row
    VraagIndex = Int((TotaalVragen - 1 + 1) * Rnd + 1)
    Worksheets(WERKBLAD_SPEL).Range("A1").Value = Worksheets(WERKBLAD_VRAGEN).Range("A" & Vraagnummer + 1).Value
    Vraagnummer = Vraagnummer + 1
    Worksheets(WERKBLAD_SPEL).Range(CEL_HUIDIGE_VRAAG).Value = Vraagnummer
   ' Bereid de boodschap voor de gebruiker voor
    Dim UserMessage As String
    Select Case Vraagnummer
        Case 1
            UserMessage = "We gaan beginnen!" & vbNewLine & "Hier komt vraag nummer één!"
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
        Case 3, 6, 9, 12
            UserMessage = BerichtTeksten(WillekeurigeIndex) & Vraagnummer & ". " & vbNewLine & vbNewLine & "Als je de volgende vraag goed beantwoordt, win je 10 seconden!"
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
        Case 15
            UserMessage = "We zijn alweer toegekomen aan de laatste vraag van deze ronde!" & vbNewLine & vbNewLine & "Ook deze vraag is goed voor 10 seconden!"
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
        Case Else
            UserMessage = BerichtTeksten(WillekeurigeIndex) & Vraagnummer & "."
            UserForm3.Label1.Caption = UserMessage
            UserForm3.Show
    End Select
    ' Reset de resterende tijd naar 20 seconden
    resterendetijd = 30
    
    ' Stop de timer indien die al loopt
    Call Aftellen_369

    ' Toon de vraag na het starten van de timer en na het sluiten van de MsgBox
    With ActiveSheet.Shapes("Vraag369").TextFrame2.TextRange.Font.Fill
        .Transparency = 0
    End With
End Sub

Sub StartSpel_369()
    ' Reset de timer
    resterendetijd = 30

    ' Reset de score
    score = 60

    ' Reset het vraagnummer
    Vraagnummer = 0

    ' Update de worksheet
    With Worksheets(WERKBLAD_SPEL)
        .Range(CEL_RESTERENDE_TIJD).Value = resterendetijd
        .Range(CEL_SCORE).Value = score
    End With
    
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
Antwoordform369.Show
End Sub

Sub SorteerVragenWillekeurig()

    Dim WerkBereik As Range
    Dim WerkArray() As Variant
    Dim TempArray() As Variant
    Dim TempVariant As Variant
    Dim i As Long, j As Long
    
   ' Definieer het bereik dat gesorteerd moet worden
    Set WerkBereik = ThisWorkbook.Worksheets(WERKBLAD_VRAGEN).Range("A1:B349")

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
    Dim HuidigeVraagnummer As Integer
    Dim resterendetijd As Integer
    Dim antwoord As String
    Dim CorrecteAntwoorden As String
    Dim CorrecteAntwoordArray() As String
    Dim i As Integer
    Dim correctGevonden As Boolean
    Dim correctIndex As Integer
    Call StopTijd_369
    With Sheets(WERKBLAD_SPEL)
        HuidigeVraagnummer = .Range(CEL_HUIDIGE_VRAAG).Value
    End With

        ' Haal het correcte antwoord op
        CorrecteAntwoorden = Sheets(WERKBLAD_VRAGEN).Cells(HuidigeVraagnummer, 2).Value
 If InStr(CorrecteAntwoorden, ";") > 0 Then
    antwoord = Split(CorrecteAntwoorden, ";")(0)
Else
    antwoord = CorrecteAntwoorden
    End If
With PasVraag369
    .Label1.Caption = "Helaas!" & vbNewLine & "Het juiste antwoord was:" & vbNewLine & vbNewLine & antwoord
    .Show
End With
    If HuidigeVraagnummer >= 15 Then
        formna369.Show
        Exit Sub
    End If
    Call VolgendeVraag
End Sub

Sub Hide()
With ActiveSheet.Shapes("Vraag369").TextFrame2.TextRange.Font.Fill
    .Transparency = 1
End With
End Sub
Sub Show()
With ActiveSheet.Shapes("Vraag369").TextFrame2.TextRange.Font.Fill
    .Transparency = 0
End With
End Sub

Sub KopieerWaarde()
    Dim huidigWerkblad As Worksheet
    Dim doelWerkblad As Worksheet

    ' Stel het huidige werkblad en het doelwerkblad in
    Set huidigWerkblad = ActiveSheet
    Set doelWerkblad = ThisWorkbook.Sheets("Open Deur")

    ' Kopieer de waarde van cel D1 van het huidige werkblad naar cel B1 van het doelwerkblad
    doelWerkblad.Range("B1").Value = huidigWerkblad.Range("D1").Value
End Sub
