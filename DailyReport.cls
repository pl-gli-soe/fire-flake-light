VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' obiekt uruchomiony z tej klasy jest zagniezdzony jako komponent obeiktu klasy FireFlakeLight
' czyli mamy tutaj do czynienia z druga warstwa pracy makra fire flake
' ten oto obiekt w ktorym teraz sie znajdujemy to DailyReport lapiacy metody po interfejsie IReport
' jego glownym komponentami jest kolekcja iteracji oraz uchwyt pojdedynczej iteracji aby latwiej pracowalo sie w petlach
' for each zeby obiekt od razu byl zdefiniowany
' wazna sprawa jest rowniez to aby pamietac ze kolekcja iteracji ma kolejne swoje elementy jako IIteration dzieki temy interfejsowi
' mozna chwytac dane z daily, hourly i weekly bez najmniejszego bolu.


Implements IReport

Private iter As IIteration
Private itersCollection As Collection
Private ffld As IFireFlakeLayout
Private dc As IDynamicColors


Private Sub Class_Initialize()

    Set mgo_handler = Nothing
    Set iter = New DailyIteration
    Set itersCollection = New Collection
    Set ffld = New FireFlakeLayoutDaily
    Set dc = New DailyDynamicColors
End Sub

Private Sub Class_Terminate()
    Set mgo_handler = Nothing
    Set iter = Nothing
    Set itersCollection = Nothing
    Set ffld = Nothing
    Set dc = Nothing
End Sub

Private Sub IReport_clearAndStop()

End Sub

Private Sub IReport_fullyRefresh()

    dc.assignDynamicColorsrange
    dc.recalcColors
End Sub

Private Function IReport_getCollectionOfIterations() As Collection
    Set IReport_getCollectionOfIterations = Nothing
End Function

Private Sub IReport_partialyRefresh()

    dc.assignDynamicColorsrange
    dc.recalcColors
End Sub

Private Sub IReport_pause()

End Sub


' metoda connect with input list bardziej by nazwa pasowala
' jest to piewrsza metoda zaraz po uruchomieniu pre konstrukora
' mozna powiedziec ze to jest ten wlasciwy konstruktor
' mysle ze warto z mojej strony pomyslec nad pewna konwencja
' zapisu i kazdy projekt jaki zaczynam zaczac od interfejsu w stylu IObject
' ktora wymuszac bedzie uruchomieniu swoistego dynamicznego konstrktora jednak z drugiej strony
' jak stowrzyc interfejs w vba ktory ma dynamiczna ilosc argumentow :D
Private Sub IReport_repInit(zakres As Range, limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, rqm_limit As Date)



    ' nalezy podkreslic ze wszystkie te dane ktore sa do siebie przypisywane to w dalszym ciagu zonglowanie danych offline
    ' czy nie ma jeszcze wartosci dodanej
    ' jest to tylko przygotowanie kolekcji pod dane z mgo - zeby bylo spojnie

    Dim r As Range
    Set itersCollection = Nothing
    Set itersCollection = New Collection
    
    For Each r In zakres
    
        ' ten prosty warunek rozwiazuje problem
        ' schowanych wierszy przez filtr
        If Not r.EntireRow.Hidden Then
    
            Set iter = Nothing
            Set iter = New DailyIteration
            iter.connectFireFlakeLayout ffld
            
            
            ' do dnia ostatniego pazdziernika 2014 rozwiazanie krzyzowalo sie pomiedzy komponentami obiektowki co nieco myli podczas realizacji
            ' i wracania poraz kolejny do implementacji i nie wiedziec co sie dzieje do okola
            
            ' zatem co by nie meczyc tematu za bardzo nalezy dodac suba w implementacji iterfejsu co robi czesciowe uzupelnianie danych
            ' w obiekcie typu IIteration
            ' a jest nim:
            ' sub powstaly dnia 31 oct 2014: czesciowyFillNaTejIteracji jako komponenet subowy obiekty typu iiteracja
            ' ======================================
            iter.czesciowyFillNaTejIteracji r, lt, st, limit, rqm_limit
            
            
            ' dodajemy szczatkowo iteracje tylko wejsciowymi danymi
            ' praca dalsza na juz istniejacych obiektach - z kwesita po kolei dodawania kolejno rqm, pus/asn / balance
            itersCollection.Add iter
            
            Set iter = Nothing
        End If
    Next r
    
    

End Sub



' metoda najwazniejsza jesli chodzi o ffa
' czy run calego raportu
Private Sub IReport_repRun(imgo As MGO)


    ' no i teraz lecimy dalej
    ' chodzi teraz o wypelnienie reszty reszta
    ' mamy niepelny fill (wlasciwie to wypelniony nawet czesciowo tylko commondata co by dalej pojsc w ogole)
    ' na kolekcjach ale jest gotowe rusztowanie - czeka puste kolejne itemy rqm intransit i ebal.
    ' tj. ilosc itemow oraz przygotowane puste kontenery na dane z mgo.
    
    
    
    
    ' jednak warunkiem podstawowym, ktory w ogole pozwala na uruchomienie raportu jest fakt ze w naszej kolekcji w ogole wystepuje jakis item
    ' dzieki temu ze wykonalismy rusztowanie w repInit
    ' w repRun nie musimy poslugiwac sie juz tak duza iloscia
    If itersCollection.COUNT > 0 Then
        If MGO_active(imgo) Then
        
        
            ' to je ino handler
            ' nie trzeba robic zadnego new DailyIteration
            ' poniewaz koniec koncow itersCollection zawiera juz szczatkowe info w sobie
            ' teraz juz trzeba tylko uzupelnic brakujace dane w juz istniejacych obiektach o odpowiednich typach.
            ' ograniczamy poza tym argumentacje w subie z 5 do jednego :)
            Dim mitem As IIteration
        
        
        
            ' i teraz male przesuniecie tej metody na sama gore
            ' aby latwiej znalezc pola ktore iteresuja pierwsza formule ebal ffa
            ffld.buildLayoutForCommonData
            
            iledni = 0
        
            For Each mitem In itersCollection
            
                ' jestesmy w srodku iteracji kazdego elementu jednorazowo
                ' mozemy teraz zrobic scenariusz
                ' tutaj dzieje sie kazdy element zwiazany z mgo
                mitem.createFullIteration imgo
                ' po uzupelnieniu danych nakladamy layout co by dobrze wygladalo
                ' i teraz tak - zakladam ze ta metoda ponizej oprocz samego koloru wyplenia rowniez
                'odpowiednio dane wg ustalonego typu runu czy raczej layotutu
                ' teraz jest ok
                ' Debug.Print mitem.getItemsCollection.COUNT
                mitem.addColorFlavour
            
            Next mitem
            ffld.buildLayoutForDatesLabels itersCollection
            
            dc.assignDynamicColorsrange
            dc.recalcColors
        End If
    End If

End Sub

Private Sub IReport_setCollectionOfIterations(coi As Collection)
    Set itersCollection = coi
End Sub
