VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakeLayoutDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' jest to zdefiniowanie daily layout
' prosto z interfejsu fire flake layout
'
'
' sprawa narazie zostanie przeze mnie rozpatrzona dosyc niespoknie tj.
' inne klasy bede traktowal dla new report
' i inne klasy bede traktowal dla continue report ...

' this jest typowo pod stworzenie nowego raportu daily
' szczegolnie widac to po metodzie pseudo konstruktora
' gdzie jawnie odnosze sie do konkretnych adresow raportu
' ktore maja sens bytu tylko dla szczegolnego przypadku stworzenia nowego raportu typu daily



Implements IFireFlakeLayout

Private tl As TheLayout
Private commonDataColumns As Range
Private toShowFromPop As Range


Private komentarzHandler As Komentarz

Private poczatekRaportuDaily As Range
Private poczatekCovLabels As Range
Private ostatniElementCommonDataTejIteracji As Range
Private pierwszyElementCoverageTejIteracji As Range


' komorka odniesienia - tak samo jak selekcja
' tyle ze nie bedzie ona jawna i pracowac bedzie niezaleznie od widoku dla usera
Private baza As Range
' sciezka routujaca
Private reportRoute As Range
Private weekNumOnTop As Range

' flaga pod aktywny arkusz po initLayout
Private sh As Worksheet


' colors
Private greyColor As Variant
Private yellowColor As Variant
Private blackColor As Variant
Private redColor As Variant
Private purple1Color As Variant


' flags for columns




Private Sub Class_Initialize()
    Set tl = New TheLayout
    
    Set komentarzHandler = New Komentarz
    
    
    greyColor = ThisWorkbook.Sheets("register").Range("grey").Interior.Color
    yellowColor = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
    blackColor = ThisWorkbook.Sheets("register").Range("black").Interior.Color
    redColor = ThisWorkbook.Sheets("register").Range("red").Interior.Color
    purple1Color = ThisWorkbook.Sheets("register").Range("purple1").Interior.Color
    
    
    
    Set commonDataColumns = ThisWorkbook.Sheets("register").Range("commonDataColumns")
    Set toShowFromPop = ThisWorkbook.Sheets("register").Range("D2:D26")
    Set weekNumOnTop = ThisWorkbook.Sheets("register").Range("weekNumOnTop")
    
    
    
    ' podliczamy z registru ile pol dodajemy do info ogolnych
    commonDataColumns = 0
    For Each i In toShowFromPop
        If i = "x" Then
            commonDataColumns = commonDataColumns + 1
        End If
    Next i
    
    'dodatkowo trzeba wziazc pod uwage:
    ' plt
    ' pn
    ' comment 1
    ' comment 2
    ' fst runout
    
    ' czyli dodatkowo 5 kolumn
    commonDataColumns = commonDataColumns + 5
    
    tl.initLayout
    ' a zostawie historyczne ze po initLayout moja jedyna flaga jest active worksheet
    Set sh = ActiveSheet
    
    Set reportRoute = sh.Range("a1")
    reportRoute = "REPORT;DAILY;LIST;" & CStr(Now)
    
    Set poczatekRaportuDaily = sh.Range("b4")
    ' MsgBox commonDataColumns
End Sub

Private Sub Class_Terminate()
    Set tl = Nothing
    Set komentarzHandler = Nothing
End Sub

Private Sub IFireFlakeLayout_autoFitContent()
    
    
    Range(poczatekRaportuDaily, poczatekRaportuDaily.End(xlToRight)).EntireColumn.AutoFit
End Sub

Private Sub IFireFlakeLayout_buildIterationLayoutForCommonData(ic As Collection, mcd As CommonData, notYetReceivedWithPastEdaCollection As Collection)
    
    ' ta procedura juz pracuje w srodku petli na kazdej iteracji po kolei
    ' na szczescie mcd zawiera mase danych dzieki czemu nie powinienem sie pogubic
    ' no ok jednak ta implementacja opiera sie na tym ze chce wygenerowac liste od nowa (forma listy)
    
    
    If mcd.layoutType = LIST_LAYOUT Then
        budujemyPodIteracjeCommonDataListLayout ic, mcd, notYetReceivedWithPastEdaCollection
    ElseIf mcd.layoutType = BOX_LAYOUT Then
    ElseIf mcd.layoutType = COV_LAYOUT Then
    End If
       
End Sub




Public Sub budujemyPodIteracjeCommonDataListLayout(ic As Collection, mcd As CommonData, notYetReceivedWithPastEdaCollection As Collection)

    Set baza = sh.Cells(mcd.rowFlag, 2)
    baza = mcd.pn
    baza.Font.size = baza.Font.size + 1
    tl.fillThinFrame baza, CDbl(greyColor)
    tl.backColor baza, CDbl(blackColor)
    baza.Font.Color = CDbl(yellowColor)
    
    baza.Offset(0, 1) = mcd.plt
    tl.fillThinFrame baza.Offset(0, 1), CDbl(greyColor)
    tl.backColor baza.Offset(0, 1), CDbl(blackColor)
    baza.Offset(0, 1).Font.Color = CDbl(yellowColor)
    
    Dim i As Range
    flaga = 2
    For Each i In toShowFromPop
        If i = "x" Then
            run CStr(i.Offset(0, -1)), mcd, baza.Offset(0, flaga)
            tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
            
            flaga = flaga + 1
        End If
    Next i
    
    ' czyli wlasciwie przesuniecie o 5 od ostniego elementu z commondata z mgo
    ' elementow posrednich jest w sumie 5
    
    ' first runout
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
    baza.Offset(0, flaga).FormulaR1C1 = "=firstRunout()"
    flaga = flaga + 1
    
    ' cmnt 1
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
    baza.Offset(0, flaga) = ThisWorkbook.Sheets("input").Cells(mcd.rowFlag - 3, 3)
    flaga = flaga + 1
    
    ' cmnt 2
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
    baza.Offset(0, flaga) = ThisWorkbook.Sheets("input").Cells(mcd.rowFlag - 3, 4)
    flaga = flaga + 1
    
    ' not yet rec
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
    baza.Offset(0, flaga) = calcTransQty(notYetReceivedWithPastEdaCollection)
    komentarzHandler.assignTransitsIntoComment baza.Offset(0, flaga), notYetReceivedWithPastEdaCollection
    flaga = flaga + 1
    
    ' past due rqm
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(greyColor)
    baza.Offset(0, flaga) = calcPastDueRqm(ic)
    
    Set ostatniElementCommonDataTejIteracji = baza.Offset(0, flaga)
End Sub


Private Function calcPastDueRqm(ic As Collection) As Long
    calcPastDueRqm = 0
    
    
    Dim di As IItem
    For Each di In ic
    
    
        ' pierwszy warunek ktory unika danych przeszlych
        ' lecimy tylko i wylacznie wtedy gdy pierwszy dzien kowerydza to dzien conajmniej dzisiejszy
        ' pozostale dane zostana wrzucone w common data :)
        If Format(di.getDate(), "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
            calcPastDueRqm = calcPastDueRqm + CLng(di.getRqm)
        End If
    Next di
    
End Function


Private Sub IFireFlakeLayout_buildIterationLayoutForCoverage(ic As Collection, mcd As CommonData)


    

    Set pierwszyElementCoverageTejIteracji = ostatniElementCommonDataTejIteracji.Offset(0, 1)
    
    
    ' ic to kolekcja itemow w tym przypadku dailyitemow
    If mcd.layoutType = LIST_LAYOUT Then
    
    
        ' jestesmy w miejscu budowania listowego fireflake
        ' troche jest niejednoznacznosci w okreslaniu i przejmowaniu
        ' parametrow wejsciowych co i jak poniewaz
        ' na samym starcie definiuje jaki rodzaj raportu chce generowac i ledwie moge przeniesc parametry okreslajace to
        ' no ale dobra, niech strace - zalozmy ze wlasnie tak ten projekt powinien wygladac - zawsze przeciez moge sie cofnac w rozwoju
        ' i tak wykorzystac w innej formie juz napisana implementacje tym bardziej ze forma i tak potrzebuje i tak tych samych
        ' parametrow wejsciowych i nie jest az tak istotne gdzie one sie wlasciwie znajduja - chodzi bardziej o spojnosc
        ' co aby jesli kiedys wroce do kodu tego
        ' to zebym potrafil to jakos ogarnac
        
        
        budowanieKowerydzaDniowegoNaLiscie ic, mcd, pierwszyElementCoverageTejIteracji
        
    ElseIf mcd.layoutType = BOX_LAYOUT Then
    ElseIf mcd.layoutType = COV_LAYOUT Then
    End If
    
    

End Sub

Private Sub budowanieKowerydzaDniowegoNaLiscie(ic As Collection, mcd As CommonData, b As Range)


    ' ok zatem b as range
    ' to pierwszy element na ktory ja jako tako mam zwrocic uwage budujac sobie dane kowerydzowe
    
    Dim pastDueRqmLbl As Range
    Set pastDueRqmLbl = pierwszyElementCoverageTejIteracji.Parent.Cells(4, pierwszyElementCoverageTejIteracji.Column - 1)
    
    ileItemow = ic.COUNT
    Dim addBackground As Boolean, firstTime As Boolean
    addBackground = False
    firstTime = True
    Dim di As IItem
    For Each di In ic
    
    
        ' pierwszy warunek ktory unika danych przeszlych
        ' lecimy tylko i wylacznie wtedy gdy pierwszy dzien kowerydza to dzien conajmniej dzisiejszy
        ' pozostale dane zostana wrzucone w common data :)
        If Format(di.getDate(), "yyyy-mm-dd") >= Format(Now, "yyyy-mm-dd") Then
        
            b = di.getRqm
            tl.fillSolidFrame b, CDbl(greyColor)
            b.Offset(0, 1) = calcTransQty(di.getInTransits())
            komentarzHandler.assignTransitsIntoComment b.Offset(0, 1), di.getInTransits()
            
            tl.fillSolidFrame b.Offset(0, 1), CDbl(greyColor)
            
            If Not firstTime Then
            
                ' policz ile przesuwamy sie do tylu aby pobrac CBAL, MISC, NOT yet rec, past due rqms
            
                b.Offset(0, 2).FormulaR1C1 = "=RC[-3]+RC[-1]-RC[-2]"
            ElseIf firstTime Then
            
                ' -5 not yet rec
                ' -4 past due rqm (sunday)
                ' -1 trans
                ' -2 rqm
                ' to calc misc & cbal
                b.Offset(0, 2).FormulaR1C1 = "=" & CStr(offsetFor("CBAL", pastDueRqmLbl)) & "-" & CStr(offsetFor("MISC", pastDueRqmLbl)) & "-RC[-3]+RC[-4]+RC[-1]-RC[-2]"
                firstTime = False
            End If
            
            
            If addBackground Then
                'For x = 0 To 1
                '    tl.backColor b.Offset(0, x), CDbl(blackColor)
                '    b.Offset(0, x).Font.Color = CDbl(yellowColor)
                'Next x
                
                tl.backColor b.Offset(0, 2), CDbl(yellowColor)
                b.Offset(0, 2).Font.Color = CDbl(blackColor)
                b.Offset(0, 2).Font.Bold = True
                
                tl.fillSolidFrame b.Offset(0, 2), CDbl(greyColor)
            Else
                tl.backColor b.Offset(0, 2), CDbl(greyColor)
                b.Offset(0, 2).Font.Color = CDbl(yellowColor)
                b.Offset(0, 2).Font.Bold = True
                
                tl.fillSolidFrame b.Offset(0, 2), CDbl(yellowColor)
            End If
            
            
            addBackground = Not addBackground
            Set b = b.Offset(0, 3)
        End If
        
    Next di

    

End Sub

Private Function offsetFor(cm As String, r As Range) As String
    ' na wypadek, gdyby cos poszlo nie tak - nie przesuwamy w ogole :D
    offsetFor = "0"
    
    ' past due na dziendobry jest przesuniety o 3
    starter = 3
    
    Dim tmp As Range
    Set tmp = r
        
    ' Debug.Print tmp - ze to niby od Past due
    Do
    
        If CStr(tmp) = cm Then
            offsetFor = "RC[-" & CStr(starter) & "]"
            Exit Function
        End If
    
        starter = starter + 1
        Set tmp = tmp.Offset(0, -1)
    Loop While tmp.Row > 1
    
    
End Function

Private Function calcTransQty(c As Collection)
    
    calcTransQty = 0
    Dim t As ITransit
    If c.COUNT > 0 Then
        For Each t In c
            calcTransQty = calcTransQty + CLng(t.mQty)
        Next t
    End If
End Function



Private Sub IFireFlakeLayout_buildLayoutForCommonData()


    ' LABELS -> OUTSIDE FROM ITERATION STUFF

    ' procedura ta ma pobierac dane z componentu common data i budowac wstepnie zarys danych na ffl
    ' teraz dobrze by bylo formualrz odpowiedni sterujacy
    ' jako poczatek
    Set baza = poczatekRaportuDaily
    
    
    baza = "Part #"
    baza.Offset(0, 1) = "Plant"
    
    
    Dim i As Range
    flaga = 2
    For Each i In toShowFromPop
        If i = "x" Then
        
            ' ten right jest po usuniecie literki s z poczatku :)
            baza.Offset(0, flaga) = Right(i.Offset(0, -1), Len(i.Offset(0, -1)) - 1)
            flaga = flaga + 1
        End If
    Next i
    
    baza.Offset(0, flaga) = "First Runout"
    baza.Offset(0, flaga + 1) = "Comment 1"
    baza.Offset(0, flaga + 2) = "Comment 2"
    baza.Offset(0, flaga + 3) = "Not yet received"
    baza.Offset(0, flaga + 4) = "Past due rqms"
    
    
    Set baza = Range(poczatekRaportuDaily, poczatekRaportuDaily.End(xlToRight))
    tl.changeTxtOrientation baza, 90
    tl.fillSolidGridLines baza, ThisWorkbook.Sheets("register").Range("grey").Interior.Color
    
    Dim r As Range
    
    For Each r In baza
        If r.Column < 4 _
            Or r = "First Runout" _
            Or r = "MISC" _
            Or r = "CBAL" _
                Then
            r.Interior.Color = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
            r.Font.Color = ThisWorkbook.Sheets("register").Range("black").Interior.Color
            r.Font.size = r.Font.size + 1
        Else
            r.Interior.Color = ThisWorkbook.Sheets("register").Range("grey").Interior.Color
            r.Font.Color = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
        End If
        
        r.Font.size = r.Font.size + 2
        
        
    Next r
    
    baza.EntireColumn.AutoFit
    Rows("4:4").RowHeight = 120
    
    
    
End Sub

Private Sub IFireFlakeLayout_buildLayoutForDatesLabels(c As Collection)


    ' c to glowna kolekcja - tak ta potezna zaiwerajaca wszystkie pnsy
    iledni = 0
    
    Dim it As IIteration
    For Each it In c
        If it.getItemsCollection.COUNT > iledni Then
            iledni = it.getItemsCollection.COUNT
        End If
    Next it
    
    Dim r As Range
    Set poczatekCovLabels = poczatekRaportuDaily.End(xlToRight).Offset(0, 1)
    Set r = poczatekCovLabels
    Dim addBackground As Boolean
    addBackground = False
    
    For x = 1 To iledni
    
        r = "Requirements"
        r.Offset(0, 1) = "In Transit"
        r.Offset(0, 2) = "Ending Balance"
        
        tl.mergeThreeCellsForDate r.Row - 1, r.Column
        tl.mergeThreeCellsForDate r.Row - 2, r.Column
        r.Offset(-1, 0) = CStr(CDate(CDate(Format(Now, "yyyy-mm-dd")) - 1 + x))
        ' poczatek weeknum w poniedzialek tj. drugi arg == 2
        If CLng(weekNumOnTop) = 1 Then
            r.Offset(-2, 0) = Application.WorksheetFunction.WeekNum(CDate(CDate(Format(Now, "yyyy-mm-dd")) - 1 + x), 2)
        End If
        
        r.Offset(-1, 0).Font.Bold = True
        r.Offset(-2, 0).Font.Bold = True
        
        
        
        
        ' r.Offset(-1, 0)
        
        For y = 0 To 2
            tl.changeTxtOrientation r.Offset(0, y), 90
        Next y
        
        
        If Not addBackground Then
            r.Font.Color = CDbl(greyColor)
            tl.fillSolidFrame r, CDbl(greyColor)
            
            r.Offset(0, 1).Font.Color = CDbl(greyColor)
            tl.fillSolidFrame r.Offset(0, 1), CDbl(greyColor)
            
            r.Offset(0, 2).Font.Color = CDbl(yellowColor)
            'r.Offset(0, 2).Font.Bold = True
            tl.fillSolidFrame r.Offset(0, 2), CDbl(yellowColor)
            tl.backColor r.Offset(0, 2), CDbl(greyColor)
            
            
            r.Offset(-1, 0).Font.Color = CDbl(yellowColor)
            r.Offset(-2, 0).Font.Color = CDbl(yellowColor)
            
            tl.backColor r.Offset(-1, 0), CDbl(greyColor)
            tl.backColor r.Offset(-2, 0), CDbl(greyColor)
            
            tl.fillSolidFrame Range(r.Offset(-1, 0), r.Offset(-1, 2)), CDbl(greyColor)
            tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(greyColor)
            
        Else
        
            r.Font.Color = CDbl(greyColor)
            tl.fillSolidFrame r, CDbl(greyColor)
            
            r.Offset(0, 1).Font.Color = CDbl(greyColor)
            tl.fillSolidFrame r.Offset(0, 1), CDbl(greyColor)
            r.Offset(0, 2).Font.Color = CDbl(blackColor)
            'r.Offset(0, 2).Font.Bold = True
            tl.fillSolidFrame r.Offset(0, 2), CDbl(greyColor)
            tl.backColor r.Offset(0, 2), CDbl(yellowColor)
            
            r.Offset(-1, 0).Font.Color = CDbl(blackColor)
            r.Offset(-2, 0).Font.Color = CDbl(blackColor)
            
            tl.backColor r.Offset(-1, 0), CDbl(yellowColor)
            tl.backColor r.Offset(-2, 0), CDbl(yellowColor)
            
            tl.fillSolidFrame Range(r.Offset(-1, 0), r.Offset(-1, 2)), CDbl(yellowColor)
            tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(yellowColor)
            
        End If
        
        If Application.WorksheetFunction.Weekday(r.Offset(-1, 0), 2) = 6 Or Application.WorksheetFunction.Weekday(r.Offset(-1, 0), 2) = 7 Then
            tl.backColor r.Offset(-1, 0), CDbl(blackColor)
            tl.backColor r.Offset(-2, 0), CDbl(blackColor)
            
            r.Offset(-1, 0).Font.Color = CDbl(yellowColor)
            r.Offset(-2, 0).Font.Color = CDbl(yellowColor)
            
            tl.fillThinGridLines Range(r.Offset(-2, 0), r.Offset(-1, 2)), CDbl(blackColor)
            tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-1, 2)), CDbl(yellowColor)
            ' tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(yellowColor)
        End If
        
        Set r = r.Offset(0, 3)
        addBackground = Not addBackground
    Next x
End Sub
