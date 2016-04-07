VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakeLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rep As IReport
Private imgo As MGO

' komorka a2 w arkuszu input
Private begOfInputList As Range
' koniec listy wejsciowej za pomoca metody end xlDown
Private endOfInputList As Range

' caly zakres danych liczac tylko z perspektywy
' pierwszej tylko kolumny
Private firstColumnScopeOnInitWorksheet As Range


' alias do komorki w arkuszu register
Private HOW_MANY_DAYS_FOR_PPUS0 As Range

' on the run property
Private ontherun As Range


' obiekt gragiczny reprezentujacy status
Private sh As StatusHandler


' jak pieknie widac na zalaczonym obrazku - pseudo konstruktor klasy ff light
' sluzy tylko i wylacznie do okreslenia zakresow wejsciowych
' dosyc zabawne myslac z perspektywy nazwy tej metody
' ale to raczej wina ze nie ma mozliwosci argumentowania konstruktora w vba
' jesli wrocisz poprzeczke wyzej do ModuleMain to zobaczysz ze new FireFlakeLight jest metoda grzecznosciowa
' i glownym elementem uruchomieniowym jest run na odpowiednim raporcie jak na przyklad
' Public Sub runDaily(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Date)
' ^^^^^^^^^^^^^^^^^^^^^
' mozna rzec ze to jest nasz konstruktor juz bardziej przyjmuje wejscie od usera i na tej bazie mozna juz pracowac
' koniec koncow taka a nie inna logika dzialania tego tutaj ponizej pseudo konstruktora jest miejsce do initu
' to uzupelniam nim to co nie wymaga customowej akcji ze strony usera.
Private Sub Class_Initialize()
    Set rep = Nothing
    Set imgo = New MGO
    
    Set sh = New StatusHandler
    sh.init_statusbar 100
    sh.show
    
    
    Set HOW_MANY_DAYS_FOR_PPUS0 = ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0")
    Set ontherun = ThisWorkbook.Sheets("register").Range("ontherun")
    ontherun = 1
    
    Set begOfInputList = ThisWorkbook.Sheets("input").Range("a2")
    If begOfInputList.Offset(1, 0) <> "" Then
        Set endOfInputList = begOfInputList.End(xlDown)
        Set firstColumnScopeOnInitWorksheet = Range(begOfInputList, endOfInputList)
    Else
        Set endOfInputList = begOfInputList
        Set firstColumnScopeOnInitWorksheet = begOfInputList
    End If
End Sub

Private Sub Class_Terminate()
    Set rep = Nothing
    Set imgo = Nothing
    ontherun = 0
    
    
    sh.hide
    Set sh = Nothing
    MsgBox "ready!"
End Sub




' run on std reports
' dodatakowo trzeba bedzie wziazc pod uwage ze raczej nie bede chcial zadnych wiecej dodatkow
' jesli chodzi o ta odslone ff'a chodzi o to aby zyl on w zgodzie z jego wlasnym nazewnictwem tj. LIGHT!
' problem pojawia sie i tak na konfigurowalnosci fire flakea mam 3 metody ktore niezaleznie pracuja od
' siebie chyba zeby stworzyc im wspolne komponenty - i raczej w ta strone bede szedl
' poniewaz kazdy z tych runow koniec koncow bedzie mial wiele cech wspolnych jak chocby algorytm wyciagania
' asn/ pus z po400 czy tez ppum0 - no i tez koniec weryfikacja danych na podstawie ekranu histori
' musze ponownie sprawdzic na jakiej zasadzie powinno to dzialac
Public Sub runDaily(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Date)

    ' zrobimy najpierw schemat w komentarzu
    ' zaczynamy od tego jaki to jest rodzaj startu
    
    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        Set rep = New DailyListReport
        ' metoda connect with input list bardziej by nazwa pasowala
        ' teraz wazna sprawa dzialanie calych iteracji opierac sie badzie na obiektach chwytajacych
        ' i tutaj uwaga juz na startcie to jest na metodzie initujacej uruchamiam kolekcje wszsytkich iteracji
        ' oczywiscie na poczatku wypelnione zostana tylko dane wejsciowe w postaci pn & plt itd
        ' cala reszta pozostanie nothing dopoty dopoki nie uruchomie nastpenych metod.
        
        
        ' trzy glowne metody pracy na kazdym raporcie light fire flake
        ' 1. init - przygotowuje rusztowanie pod wszystkie dane
        
        ' 2. connect with mgo chyba nie wymaga szerwszego komentarza
        '   uruchamia nam przy okazji wszelkie pomocniki obiektowe ktore pozwalaja nam zapomniec o pracy na adresach ale bezposrednio
        '   odnosiemy sie nazwami do pol sesji mgo
        
        ' 3. repRun jest to juz gotowa metoda ktora czerpie z dwoch poprzednich czesciowo wypelniona kolekcja pustymi obiektami zostaje zapelniona wartosciami
        '   z sesji mgo.
        
        rep.repInit firstColumnScopeOnInitWorksheet, limit, lt, st, CDate(daily_rqm_limit)
        rep.repRun imgo, sh
        
        
        
        
        ' =======================================================================================================
        
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        ' a wiec z tego co rozumiem continue jest realizowany
        ' poza scopem tej procedury!
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If
    
    
    rep.fakeDataChangeToRecalcFirstRunoutFormula
    ' rep.adjustLayoutAfterAllJobs

End Sub

Public Sub runHourly(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Integer)

    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        
        
        ' =======================================================================================================
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If

End Sub

Public Sub runWeekly(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Integer)

    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        
        
        ' =======================================================================================================
        
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If

End Sub

Public Sub continueBrokenReport()
    ' nie ptorzebuje zadnych dodatkowych danych
    ' sprawdz czy lista wejsciowa dalej jest zsynchr razem ze zepsutym reportem
    
    ' tutaj warunek dla st w ogole nie jest potrzebny :D
    ' poniewaz procedura i tak nie przyjmuje zadnego argumentu :D
    ' lecimy od razu z algorytm odzyskiwania informacji jak dalej wykonac niedokonczony raport ff
    ' =======================================================================================================
    
    
    
    
    
    ' =======================================================================================================
End Sub
