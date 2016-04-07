VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakeLayoutDailyList"
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


' jest tez tutaj mala niescislosc zwiazana z tym ze wlasciwie ta klasa odpowiada tylko i wylacznie za layout listowy



Implements IFireFlakeLayout

Private tl As TheLayout
Private commonDataColumns As Range
Private toShowFromPop As Range


Private komentarzHandler As Komentarz

Private poczatekRaportuDaily As Range
Private poczatekCovLabels As Range
Private ostatniElementCommonDataTejIteracji As Range
Private pierwszyElementCoverageTejIteracji As Range


' te pola sa pelne dopiero po uzupelnieniu labelek
Private ostatni_wiersz As Range
Private ostatnia_kolumna As Range


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

' my colors on my layout
Private primaryColor As Variant
Private secondaryColor As Variant
Private weekendColor As Variant
Private minusColor As Variant
Private warningColor As Variant


' ta zmienna sprawdza
' jaki jest max wymiar w prawo
' dla wszystkich danych
' dzieki czemu mozna sie potem odnosic podczas wystawiania labelek
' czy tez wsuwania thin grid lines
Private iledni As Long


' flags for columns




Private Sub Class_Initialize()
    Set tl = New TheLayout
    
    Set komentarzHandler = New Komentarz
    
    
    greyColor = ThisWorkbook.Sheets("register").Range("grey").Interior.Color
    yellowColor = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
    blackColor = ThisWorkbook.Sheets("register").Range("black").Interior.Color
    redColor = ThisWorkbook.Sheets("register").Range("red").Interior.Color
    purple1Color = ThisWorkbook.Sheets("register").Range("purple1").Interior.Color
    
    primaryColor = ThisWorkbook.Sheets("register").Range("primary").Interior.Color
    secondaryColor = ThisWorkbook.Sheets("register").Range("secondary").Interior.Color
    weekendColor = ThisWorkbook.Sheets("register").Range("weekendColor").Interior.Color
    minusColor = ThisWorkbook.Sheets("register").Range("minus").Interior.Color
    warningColor = ThisWorkbook.Sheets("register").Range("warning").Interior.Color
    
    
    
    Set commonDataColumns = ThisWorkbook.Sheets("register").Range("commonDataColumns")
    
    ' tutaj troche kombinacja alpejska
    ' jednak ze chcemy zeby bylo dynamicznie
    ' i zeby kodu nie trzeba bylo zmieniac na poziomie kilku miejsc w rozych modulach
    ' najpierw z gory do dolu po labelkach i przejscie jednym offsetem na iksy
    Set toShowFromPop = ThisWorkbook.Sheets("register").Range("C2")
    Set toShowFromPop = ThisWorkbook.Sheets("register").Range(toShowFromPop, toShowFromPop.End(xlDown))
    Set toShowFromPop = toShowFromPop.Offset(0, 1)
    
    Set weekNumOnTop = ThisWorkbook.Sheets("register").Range("weekNumOnTop")
    iledni = 0
    
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
    
    
End Sub

Private Sub Class_Terminate()
    Set tl = Nothing
    Set komentarzHandler = Nothing
End Sub

Private Sub IFireFlakeLayout_buildLayoutTheThinLines(c As Collection)

    ' ile dni zostalo juz wpisane w poprzedniej procedurze
    ' ==================================================================
    ' IFireFlakeLayout_buildLayoutForDatesLabels
    ' ale tez nie wiem na ile bedzie to faktycznie potrzebne
    
    
    ' te przypisanie dopiero teraz ma sens poniewaz dopiero teraz dane labelek
    ' zostaly dopisane
    Set ostatnia_kolumna = poczatekRaportuDaily.End(xlToRight)
    Set ostatni_wiersz = poczatekRaportuDaily.End(xlDown)
    
    tl.fillThinGridLines sh.Range(sh.Cells(poczatekRaportuDaily.Row + 1, poczatekRaportuDaily.Column + 2), sh.Cells(ostatni_wiersz.Row, ostatnia_kolumna.Column)), CLng(primaryColor)
    
    ' ==================================================================
End Sub

Private Sub IFireFlakeLayout_closingDayLines()

    ' metoda wykonywana na samym koncu layoutowania
    ' mam juz niezbedne dane jak ostatnia kolumna i ostatni wiersz
    ' plus musze jescze wiedziec kiedy zaczynaja sie ending balance
    ' poczatekCovLabels ta zmienna to w sumie pierwszy rqm
    ' te pola sa pelne dopiero po uzupelnieniu labelek
    ' Private ostatni_wiersz As Range
    ' Private ostatnia_kolumna As Range
    
    
    Dim r As Range
    Set r = poczatekCovLabels.Offset(0, 2)
    Do
        Set r = r.Parent.Range(Cells(r.Row, r.Column), Cells(ostatni_wiersz.Row, r.Column))
        tl.fillRightThinLine r, CLng(blackColor)
        Set r = r.Item(1, 1).Offset(0, 3)
    Loop Until Trim(r) = ""
    
End Sub

Private Sub IFireFlakeLayout_finalAdjustOnColumnsWidth()
    sh.Range("a1").ColumnWidth = 1
    ' sh.Range(sh.Cells(1, poczatekRaportuDaily.Column + 2), sh.Cells(1, ostatniElementCommonDataTejIteracji.Column)).ColumnWidth = 4
    Dim r_to_adj As Range
    For Each r_to_adj In sh.Range(sh.Cells(1, poczatekRaportuDaily.Column + 2), sh.Cells(1, ostatniElementCommonDataTejIteracji.Column))
    
        r_to_adj.ColumnWidth = 6
    Next r_to_adj
    
End Sub

Private Sub IFireFlakeLayout_freezePanesAtTheEnd(ByRef lt As LAYOUT_TYPE)
    If lt = LIST_LAYOUT Then
        Range("D5").Select
        ActiveWindow.freezePanes = True
    End If
End Sub

Private Sub IFireFlakeLayout_addFilterOnListLayout()
    poczatekRaportuDaily.Parent.Range(poczatekRaportuDaily, poczatekRaportuDaily.End(xlToRight)).AutoFilter
End Sub

Private Sub IFireFlakeLayout_autoFitContent()
    
    
    Range(poczatekRaportuDaily, poczatekRaportuDaily.End(xlToRight)).EntireColumn.AutoFit
End Sub

Private Sub IFireFlakeLayout_buildIterationLayoutForCommonData(ByRef ic As Collection, ByRef mcd As CommonData, ByRef notYetReceivedWithPastEdaCollection As Collection)
    
    ' ta procedura juz pracuje w srodku petli na kazdej iteracji po kolei
    ' na szczescie mcd zawiera mase danych dzieki czemu nie powinienem sie pogubic
    ' no ok jednak ta implementacja opiera sie na tym ze chce wygenerowac liste od nowa (forma listy)
    
    
    If mcd.layoutType = LIST_LAYOUT Then
        budujemyPodIteracjeCommonDataListLayout ic, mcd, notYetReceivedWithPastEdaCollection
    ElseIf mcd.layoutType = BOX_LAYOUT Then
    ElseIf mcd.layoutType = COV_LAYOUT Then
    End If
       
End Sub




Private Sub budujemyPodIteracjeCommonDataListLayout(ByRef ic As Collection, ByRef mcd As CommonData, ByRef notYetReceivedWithPastEdaCollection As Collection)

    ' Debug.Print "budujemyPodIteracjeCommonDataListLayout beg " & CStr(Now)
    ' init ref variable
    ' this is a base from comment as well
    ' ==========================================================================
    Set baza = sh.Cells(mcd.rowFlag, 2)

    
    ' fill data from pop screen into comment in first cell (pn)
    ' ==========================================================================
    Dim k As Komentarz
    Set k = New Komentarz
    k.typeOfComment = DATA_FROM_POP
    k.makeCommentBasedOnPopData baza, mcd
    Set k = Nothing

    ' fill data in cells
    ' ==========================================================================

    
    baza = mcd.pn
    ' baza.Font.size = baza.Font.size + 1
    tl.fillThinFrame baza, CDbl(primaryColor)
    tl.backColor baza, CDbl(secondaryColor)
    baza.Font.Color = CDbl(blackColor)
    
    baza.Offset(0, 1) = mcd.plt
    tl.fillThinFrame baza.Offset(0, 1), CDbl(primaryColor)
    tl.backColor baza.Offset(0, 1), CDbl(secondaryColor)
    baza.Offset(0, 1).Font.Color = CDbl(blackColor)
    
    Dim i As Range
    flaga = 2
    For Each i In toShowFromPop
        If i = "x" Then
            run CStr(i.Offset(0, -1)), mcd, baza.Offset(0, flaga)
            tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
            
            flaga = flaga + 1
        End If
    Next i
    
    ' czyli wlasciwie przesuniecie o 5 od ostniego elementu z commondata z mgo
    ' elementow posrednich jest w sumie 5
    
    ' first runout
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
    ' RC[-8] - to nie jest zawsze prawda!
    ile_na_minus = CStr(baza.Offset(0, flaga).Column - 1)
    ' tu jest troche dupa bo statycznie przypisalem tylko 100 komorek co nie jest do konca prawda :D
    ' ile_na_plus w sumie tez sie przyda :D
    ' ile_na_plus = CStr(baza.Parent.Range("b4").End(xlToRight))
    ile_na_plus = CStr(200)
    ' bedzie te ile_na_plus treoche redundantne ale tylko o kilka zatem nie powinno robic problemu ostatecznie
    baza.Offset(0, flaga).FormulaR1C1 = "=firstRunout(RC[-" & CStr(ile_na_minus) & "]:RC[-1], RC[1]:RC[" & ile_na_plus & "])"
    flaga = flaga + 1
    
    ' cmnt 1
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
    baza.Offset(0, flaga) = ThisWorkbook.Sheets("input").Cells(mcd.rowFlag - 3, 3)
    flaga = flaga + 1
    
    ' cmnt 2
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
    baza.Offset(0, flaga) = ThisWorkbook.Sheets("input").Cells(mcd.rowFlag - 3, 4)
    flaga = flaga + 1
    
    ' not yet rec
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
    baza.Offset(0, flaga) = calcTransQty(notYetReceivedWithPastEdaCollection)
    komentarzHandler.assignTransitsIntoComment baza.Offset(0, flaga), notYetReceivedWithPastEdaCollection
    flaga = flaga + 1
    
    ' past due rqm
    tl.fillThinFrame baza.Offset(0, flaga), CDbl(primaryColor)
    baza.Offset(0, flaga) = calcPastDueRqm(ic)
    
    Set ostatniElementCommonDataTejIteracji = baza.Offset(0, flaga)
    
    
    ' Debug.Print "budujemyPodIteracjeCommonDataListLayout end " & CStr(Now)
End Sub


Private Function calcTransQty(c As Collection)
    
    calcTransQty = 0
    Dim t As ITransit
    If c.COUNT > 0 Then
        For Each t In c
            calcTransQty = calcTransQty + CLng(t.mQty)
        Next t
    End If
End Function


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


Private Sub IFireFlakeLayout_buildIterationLayoutForCoverage(ByRef ic As Collection, ByRef mcd As CommonData)

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

Private Sub budowanieKowerydzaDniowegoNaLiscie(ic As Collection, mcd As CommonData, ByRef b As Range)


    ' algorytm z 14 kwietnia na tej procedurze zajmuje 4 sekundy!
    ' [sic!]
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    
    ' ok zatem b as range
    ' to pierwszy element na ktory ja jako tako mam zwrocic uwage budujac sobie dane kowerydzowe
    
    Dim pastDueRqmLbl As Range
    Set pastDueRqmLbl = b.Parent.Cells(4, b.Column - 1)
    
    ileItemow = ic.COUNT
    Dim firstTime As Boolean
    ' addBackground = False
    firstTime = True
    
    
    ' dodatkowo przed petla aby uniknac petli w petli
    ' gdyz zawsze jest to problem dla vba
    ' wczesniej zadeklarujemy odleglosci miedyz CBAL i MISC od pierwszego EBALA
    ofst_CBAL = CStr(offsetFor("CBAL", pastDueRqmLbl))
    ofst_MISC = CStr(offsetFor("MISC", pastDueRqmLbl))
    for_first_time_formula = "=" & CStr(ofst_CBAL) & "-" & CStr(ofst_MISC) & "-RC[-3]+RC[-4]+RC[-1]-RC[-2]"
    for_regular_formula = "=RC[-3]+RC[-1]-RC[-2]"
    
    Dim di As IItem
    For Each di In ic
    
        ' pierwszy warunek ktory unika danych przeszlych
        ' lecimy tylko i wylacznie wtedy gdy pierwszy dzien kowerydza to dzien conajmniej dzisiejszy
        ' pozostale dane zostana wrzucone w common data :)
        If Format(di.getDate(), "yyyy-mm-dd") >= Format(Now, "yyyy-mm-dd") Then
        
            b = di.getRqm
            'tl.fillSolidFrame b, CDbl(primaryColor)
            ' 1 sec obsolete
            ' b.Offset(0, 1) = calcTransQty(di.getInTransits())
            If di.getSumOfQty() = 0 Then
                b.Offset(0, 1) = ""
            Else
                b.Offset(0, 1) = di.getSumOfQty()
            End If
            komentarzHandler.assignTransitsIntoComment b.Offset(0, 1), di.getInTransits()
            
            'tl.fillSolidFrame b.Offset(0, 1), CDbl(primaryColor)
            
            
            If firstTime Then
            
                ' -5 not yet rec
                ' -4 past due rqm (sunday)
                ' -1 trans
                ' -2 rqm
                ' to calc misc & cbal
                b.Offset(0, 2).FormulaR1C1 = CStr(for_first_time_formula)
                firstTime = False
            Else
            
                ' policz ile przesuwamy sie do tylu aby pobrac CBAL, MISC, NOT yet rec, past due rqms
            
                b.Offset(0, 2).FormulaR1C1 = CStr(for_regular_formula)
            End If
            
            
            b.Offset(0, 2).Font.Color = CDbl(blackColor)
            b.Offset(0, 2).Font.Bold = True
            'If addBackground Then
            '
            '    tl.backColor b.Offset(0, 2), CDbl(primaryColor)
            '    b.Offset(0, 2).Font.Color = CDbl(blackColor)
            '    b.Offset(0, 2).Font.Bold = True
            '
            '    tl.fillSolidFrame b.Offset(0, 2), CDbl(primaryColor)
            'Else
            '    tl.backColor b.Offset(0, 2), CDbl(secondaryColor)
            '    b.Offset(0, 2).Font.Color = CDbl(blackColor)
            '    b.Offset(0, 2).Font.Bold = True
            '
            '    tl.fillSolidFrame b.Offset(0, 2), CDbl(primaryColor)
            'End If
            
        
            ' addBackground = Not addBackground
            Set b = b.Offset(0, 3)
            
        End If
        
        
    Next di
    
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    
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
    tl.fillSolidGridLines baza, CLng(secondaryColor)
    
    Dim r As Range
    
    For Each r In baza
        If r.Column < 4 _
            Or r = "First Runout" _
            Or r = "MISC" _
            Or r = "CBAL" _
                Then
            r.Interior.Color = primaryColor
            r.Font.Color = blackColor
            r.Font.size = r.Font.size + 1
        Else
            r.Interior.Color = secondaryColor
            r.Font.Color = blackColor
        End If
        
        r.Font.size = r.Font.size + 2
        
        
    Next r
    
    baza.EntireColumn.AutoFit
    Rows("4:4").RowHeight = 120
    
    
    
End Sub

Private Sub IFireFlakeLayout_buildLayoutForDatesLabels(ByRef c As Collection)


    ' c to glowna kolekcja - tak ta potezna zaiwerajaca wszystkie pnsy
    ' wczesniej byla to zmienna lokalna teraz
    ' jest to pole dzieki czemu po tej procedurze inne tez beda mogly z tego korzystac
    ' iledni = 0
    
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
        r.Offset(-1, 0).HorizontalAlignment = xlCenter
        
        ' poczatek weeknum w poniedzialek tj. drugi arg == 2
        If CLng(weekNumOnTop) = 1 Then
            r.Offset(-2, 0) = Application.WorksheetFunction.WeekNum(CDate(CDate(Format(Now, "yyyy-mm-dd")) - 1 + x), 2)
            r.Offset(-2, 0).HorizontalAlignment = xlLeft
        End If
        
        r.Offset(-1, 0).Font.Bold = True
        r.Offset(-2, 0).Font.Bold = True
        
        
        
        
        ' r.Offset(-1, 0)
        
        For y = 0 To 2
            tl.changeTxtOrientation r.Offset(0, y), 90
        Next y
        
        
        If Not addBackground Then
            r.Font.Color = CDbl(blackColor)
            tl.fillSolidFrame r, CDbl(secondaryColor)
            
            r.Offset(0, 1).Font.Color = CDbl(blackColor)
            tl.fillSolidFrame r.Offset(0, 1), CDbl(secondaryColor)
            
            r.Offset(0, 2).Font.Color = CDbl(blackColor)
            'r.Offset(0, 2).Font.Bold = True
            tl.fillSolidFrame r.Offset(0, 2), CDbl(primaryColor)
            tl.backColor r.Offset(0, 2), CDbl(secondaryColor)
            
            
            r.Offset(-1, 0).Font.Color = CDbl(blackColor)
            r.Offset(-2, 0).Font.Color = CDbl(blackColor)
            
            tl.backColor r.Offset(-1, 0), CDbl(secondaryColor)
            tl.backColor r.Offset(-2, 0), CDbl(secondaryColor)
            
            tl.fillSolidFrame Range(r.Offset(-1, 0), r.Offset(-1, 2)), CDbl(secondaryColor)
            tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(secondaryColor)
            
        Else
        
            r.Font.Color = CDbl(blackColor)
            tl.fillSolidFrame r, CDbl(secondaryColor)
            
            r.Offset(0, 1).Font.Color = CDbl(blackColor)
            tl.fillSolidFrame r.Offset(0, 1), CDbl(secondaryColor)
            r.Offset(0, 2).Font.Color = CDbl(blackColor)
            'r.Offset(0, 2).Font.Bold = True
            tl.fillSolidFrame r.Offset(0, 2), CDbl(secondaryColor)
            tl.backColor r.Offset(0, 2), CDbl(primaryColor)
            
            r.Offset(-1, 0).Font.Color = CDbl(blackColor)
            r.Offset(-2, 0).Font.Color = CDbl(blackColor)
            
            tl.backColor r.Offset(-1, 0), CDbl(primaryColor)
            tl.backColor r.Offset(-2, 0), CDbl(primaryColor)
            
            tl.fillSolidFrame Range(r.Offset(-1, 0), r.Offset(-1, 2)), CDbl(primaryColor)
            tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(primaryColor)
            
        End If
        
        If Application.WorksheetFunction.Weekday(r.Offset(-1, 0), 2) = 6 Or Application.WorksheetFunction.Weekday(r.Offset(-1, 0), 2) = 7 Then
            tl.backColor r.Offset(-1, 0), CDbl(weekendColor)
            tl.backColor r.Offset(-2, 0), CDbl(weekendColor)
            
            r.Offset(-1, 0).Font.Color = CDbl(blackColor)
            r.Offset(-2, 0).Font.Color = CDbl(blackColor)
            
            'tl.fillThinGridLines Range(r.Offset(-2, 0), r.Offset(-1, 2)), CDbl(blackColor)
            'tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-1, 2)), CDbl(blackColor)
            ' tl.fillSolidFrame Range(r.Offset(-2, 0), r.Offset(-2, 2)), CDbl(yellowColor)
        End If
        
        Set r = r.Offset(0, 3)
        addBackground = Not addBackground
    Next x
End Sub


Private Sub IFireFlakeLayout_init(lt As LAYOUT_TYPE, st As START_TYPE)


    If lt = LIST_LAYOUT And st = FROM_THE_BEGINNING Then
        tl.initLayout
        ' a zostawie historyczne ze po initLayout moja jedyna flaga jest active worksheet
        Set sh = ActiveSheet
        
        Set reportRoute = sh.Range("a1")
        reportRoute = "REPORT;DAILY;LIST;" & CStr(Now) & ";RQM_LIM:" & CStr(ThisWorkbook.Sheets("register").Range("rqmLimit")) & ";PUS_LIM:" & CStr(ThisWorkbook.Sheets("register").Range("pusLimit"))
        
        Set poczatekRaportuDaily = sh.Range("b4")
        ' MsgBox commonDataColumns
    ElseIf lt = LIST_LAYOUT And st = CONTINUE_BROKEN_ONE Then
        Set sh = ActiveSheet
        Set reportRoute = sh.Range("a1")
        Set poczatekRaportuDaily = sh.Range("b4")
    End If
End Sub
