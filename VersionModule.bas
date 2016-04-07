Attribute VB_Name = "VersionModule"
' Version Module








' no to start
' modul ten powstal od czwartej generacji ffa
' w mysl o tym czego wlasciwie potrzebuje uzytkownik musze pomyslec ile faktycznie jest narzedzi potrzebnych a ile nie
' odchudzenie ffa rowniez wchodzi w gre aby chodzil odrobine szybciej
Public Sub msgbox_about(ictrl As IRibbonControl)



    ' now in transit class we have new added lines: on error resume next
    
    ' ------------------------------------------------------------
    ' On Error Resume Next
    ' t.mDeliveryDate = CDate(m.convertToDateFromMS9POP00Date(m.pMS9POP00.transEDA(Int(x))))
    ' On Error Resume Next
    ' t.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
    ' t.mNotYetReceived = True
    ' ...
    ' On Error Resume Next
    ' t.mPickupDate = CDate(m.convertToDateFromMS9POP00Date(CStr(m.pMS9POP00.transSDATE(Int(x)))))
    ' ------------------------------------------------------------
    
    version_4_03_09 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.09" & Chr(10) & _
        "Limited to daily report" & Chr(10) & _
        " - changed on calendar week logic (ISO 8601)" & _
        Chr(10)
    
    version_4_03_07 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.07" & Chr(10) & _
        "Limited to daily report" & Chr(10) & _
        " - changed on logic from continue broken report" & _
        Chr(10)
    
    
    version_4_03_06 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.06" & Chr(10) & _
        "Limited to daily report" & Chr(10) & _
        " - custom events implemented for dynamic form for controlling logic on rqm downloading" & Chr(10) & _
        " - added on error resume next in transit class to avoid errors when eda or eta is empty" & Chr(10) & _
        " - first 03.xx with fix on XFR from 02.25 important!" & Chr(10) & _
        " - dynamic wizard for new plants (unrecognized plants)" & Chr(10) & _
        Chr(10)

    
    
    version_4_03_05 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.05" & Chr(10) & _
        " - custom events implemented for dynamic form for controlling logic on rqm downloading" & Chr(10) & _
        " - " & Chr(10) & _
        " - first 03.xx with fix on XFR from 02.25 important!" & Chr(10) & _
        " - dynamic wizard for new plants (unrecognized plants)" & Chr(10) & _
        Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        ""

    ' big changes in addRqmsAndDatesIntoItems in DailyIteration class
    '4.03.04 - ma poprawe z wersji 4.02.25
    ' -------------------------------------------------------------
    version_4_03_04 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.04" & Chr(10) & _
        " - big changes in addRqmsAndDatesIntoItems in DailyIteration class" & Chr(10) & _
        " - extra private subroutines zk7pdrqm_logic & ms9pop00_logic to seperate config on component" & Chr(10) & _
        " - rqm from zk7pdrqm logic added - config 0" & Chr(10) & _
        " - rqm from ms9pop00 logic to be added - config 1" & Chr(10) & _
        " - rqm from zeros logic to be added - config 4?" & Chr(10) & _
        " - added downloading rules handler for form and dynamic events for creation form ad hoc" & Chr(10) & _
        " - first 03.xx with fix on XFR from 02.25 important!" & Chr(10) & _
        Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        " - dynamic wizard for new plants (unrecognized plants)" & Chr(10) & _
        ""


    '4.03.03 - ma poprawe z wersji 4.02.25
    ' -------------------------------------------------------------
    version_4_03_03 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.03" & Chr(10) & _
        " - input list config added" & Chr(10) & _
        " - rqm from zk7pdrqm logic to be added - config 0" & Chr(10) & _
        " - rqm from ms9pop00 logic to be added - config 1" & Chr(10) & _
        " - first 03.xx with fix on XFR from 02.25 important!" & Chr(10) & _
        " - also point on ms9ph100 tylko_na_poczatku_jest_to_zerem_zwiazane_z_ukladem_ekranu_historii! PickupHandler PH100" & Chr(10) & _
        Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        " - rules form with connections with real data" & Chr(10) & _
        ""

    '4.03.02
    ' -------------------------------------------------------------
    version_4_03_02 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.02" & Chr(10) & _
        "NEW RULES FOR COMPONENT PLANTS: " & Chr(10) & _
        " - from data downloading perspective still no value added (only zk7pdrqm screen implemented)" & Chr(10) & _
        " - add logic for plants that not exists in register list" & Chr(10) & _
        " - only actions on form added with no actions on input list" & Chr(10) & _
        Chr(10) & _
        ""

    '4.03.01 - wersja jeszcze nie udostepniona ma blad z wersji 02.24
    ' -------------------------------------------------------------
    version_4_03_01 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.03.01" & Chr(10) & _
        "NEW RULES FOR COMPONENT PLANTS: " & Chr(10) & _
        " - this is init version for fire flake light with component logic" & Chr(10) & _
        " - idea is week because we need to work inline with plans that are in the list" & Chr(10) & _
        Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        " - rules form without further connections" & Chr(10) & _
        ""

    
    '4.02.25
    ' -------------------------------------------------------------
    version_4_02_25 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.02.25" & Chr(10) & _
        "FIXED: " & Chr(10) & _
        " - XFR without EDA brokes running" & Chr(10) & _
        " - also point on ms9ph100 tylko_na_poczatku_jest_to_zerem_zwiazane_z_ukladem_ekranu_historii! PickupHandler PH100" & Chr(10) & _
        ""

    
    '4.02.24
    ' -------------------------------------------------------------
    version_4_02_24 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.02.24" & Chr(10) & _
        "FIXED: " & Chr(10) & _
        " - Broken report reimplemented & fixed, but need to have additional tests" & Chr(10) & _
        " - Simlation on continue broken report done with OK status" & Chr(10) & _
        " - Scaffold changed on primary and secondary color on dates and weeknums" & Chr(10) & _
        " - Extend text comment with POP screen data (second column in list layout - fix)" & Chr(10) & _
        " - Filtered input list - wrong cmnt 1 & cmnt 2 (fixed)" & Chr(10) & _
        "" & Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        " - to be finnaly done as stable version" & Chr(10) & _
        ""
    
    '4.02.23 od tej wersji dane beda w zmiennej globalnej
    ' -------------------------------------------------------------
    version_4_02_23 = "This is Fire Flake Light - the 4th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 4.02.23" & Chr(10) & _
        "FIXED: " & Chr(10) & _
        " - Freeze Panes Bug fixed" & Chr(10) & _
        " - Delete sheet button group added" & Chr(10) & _
        " - About Group added (for validation)" & Chr(10) & _
        "" & Chr(10) & _
        "TO BE IMPLEMENTED: " & Chr(10) & _
        " - CONTINUE BROKEN REPORT still not working for internet connection breakdowns" & Chr(10) & _
        ""
    
    MsgBox CStr(version_4_03_09)
    
End Sub

'4.02.22
' -------------------------------------------------------------
' okazuje sie ze zapomnialem dodac kolory na cbalu
' oczywiscie rowniez podtrzymuje temat z 20 i 21
' slawkowi na wersji 02.21 freeze panes odmowil posluszenstwa z losowych przyczyn :(
' dobrze by jakis test rzucic :D
'
'
'


'4.02.21
' -------------------------------------------------------------
' z wersji 4.02.20 wicaz hold na to ponizej
' szybki powrot do cebeera
' i malych niedociagniec miedzy innymi gdy nastepuje disconnection na systemie mainframe
' ff dalej ciagnie swoja logike wstawiajac jednak tylko pn i plt :D
' no i buduje wlasciwie cale rusztowanie bez bolu
' trzeba by to dopisac
'
' poprawiona logik ana ppus0 poniewaz fil zrobilem z perspektywy delivery date a nie pickup
' robilo to dziure na kilka dni miedzy ppus0 a ph100
'
' poprawiona dziwny blad podczas uruchamiania nowego ff'a od pustych raportow
' gubil jakos mozliwosc przeliczania
' prawdopodobnie podczas usuwania wszytkich raportow register gubil sie
' wsadzajac activeshhet jako input stad typ dynamic colors nie radzil sobie z tematem i na koniec pierwszego raportu
' zwracal nieokreslony blad
' od tej pory wszsytko powinno dzialac jednak gdzies tam sie czai historia zwiazana z nazewnictwem pewnie


'4.02.20
' -------------------------------------------------------------
' szybki powrot do cebeera
' i malych niedociagniec miedzy innymi gdy nastepuje disconnection na systemie mainframe
' ff dalej ciagnie swoja logike wstawiajac jednak tylko pn i plt :D
' no i buduje wlasciwie cale rusztowanie bez bolu
' trzeba by to dopisac

'4.02.19
' -------------------------------------------------------------
' w miedzy czasie pojawil sie problem w klasie dynamic colors nie wiadomo czemu range nie chce sie dopasowac
' pomimo tego ze skladnia jest cacy
' nie wiem wlasciwie cos sie stalo ze raz dziala a raz nie
' moze sie okazac ze z klasa dynamic colors beda jeszcze w przyszlosci problemy niewiadomego pochodzenia :D

'4.02.18
' -------------------------------------------------------------
' cebeer poczatek
' udalo sie zrobic prosty continue broken report
' jednak nieprawidlowo obsluguje przerwanie lacza internetowego
' musze wykonac symulacje wylaczenia internetu
' zrobienia disconnected
' zmieny aktywnosci arkusza
' kazdy ekran musi posiadac obslugo wyjatku do tego stopnia zeby ff sie zwieszal mocniej a nie ignorowal bledy


'4.02.17
' -------------------------------------------------------------
' ta wersja bedzie slodko uzbrojona w dodatek continue broken report!

' cala idea polega na kontynuowaniu zatrzymanego ff'a
' cala logika skaldac sie bedzie z dwoch akcji:
' spradzenie czy lista wejsciowa jest kompatybilna z czesciowo wypelnionym ff'em
' uruchomienie ff'a od odpowiedniego miejsca :D
' warunki do sprawdzenia
' to czy lista jest kompatyblina
' to czy lista jest swieza - to znaczy, czy dane z ff sa w miare inline z rzeczywistoscia pokazywana na MGO
' tutaj do przemyslenia jaki czas to czas danych jako obsoletowych
' Set input_rng_flag = init_sh.Range("a2")
' Set report_rng_flag = active_sh.Range("b5")
'
' CEBEER! :D

'4.02.16
' -------------------------------------------------------------
' poprawa logiki na zaciaganiu daily rqm
' plus zaczatek logiki na continue broken report

'4.02.15
' -------------------------------------------------------------
' poprwawiona implementacja first runout formula
' If rng = "" Then
'     firstRunout = "no data"
'     Exit Function
' End If
'
'
' dorzucony element sprawdzania gdyby jednak nie bylo jeszcze wpisanych kolumn ending balance bo w koncu nie jest to musem
' zawsze moze sie okazac ze wszystkie part numbery beda puste z punktu widzenia zk7pdrqm


'4.02.14
' -------------------------------------------------------------
' pickup limit zaimplementowany znow

'4.02.13
' -------------------------------------------------------------
' wersja 64 bitowa ma type mismatch to GlobalModule dalem dodatkowe deklaracje
'#If VBA7 Then
'Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As LongPtr
'#Else
'Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'#End If
' to powyzsze jednak nie wiele dalo
' plus teraz duze zmiany jesli chodzi o layout
' wyrzucilem z petli duzo wsadzilem tuz za nia
' jak malowanie thin lines
' freeze panes
' filtering etc
'
'
' zmiana z 2015-04-17
' dodatkowo przeusniecie daty first runout o jeden dzien dalej jesli w ogole nie ma runoutu
' w przypadku filtrowania i jest runout ostatniego dnia
' minusy z plusami zaczna sie mieszac :D


'4.02.12
' -------------------------------------------------------------
' zmiana pod wersji tylko w ramach bezpieczenstwa poniewaz od 12 planuje zmiany na addFlavourColor ktore moga spowodowac zawieszenie aplikacji powazne
' dodatkowo kontrola kolorow - jakas rozszerzona paleta
' ale skrajnosci ze mozna wybrac doslownie najglubszy kolor
' albo ewentualnie tak jak jest na nowym weekly coverage
' czyli kilka gotowych wyborow plus ewentualny custom - jak ktos bedzie chcial to sobie sam sobie zmieni

'4.02.11
' -------------------------------------------------------------
' okazuje sie ze wersja 10 jest prawie 2 razy wolniejsza od poprzedniej wersji
' ffh
' plus czarno zolty to jednak beznadziejny kolor
' zatem po krotce nowa paleta barw:
' zielony (rgb, 115, 175, 173)
' blado szaro zielono niebieski (rgb, 236, 234, 234)
' blady pomarancz (rgb, 217, 163, 59)
'
'
' innym ciekawym zagraniem jest wykorzystanie slownika Dictionary (dodatkowa biblioteka wiec be aware moze komus te rozwiazanie nie zadzialac)
' w ktorym bawimy sie asocjacyjnie, gdzie keyami beda nazwy transportow SID?
' dzieki temu rozwiazuje problem duplikatow gdy przejde na ms9po400 i znow bede chcial zaciagac dane ktore juz zostaly pobrane z zk7ppum0
'
'
' dodatkowa logika sprawdzajaca czy warto w ogole przechodzic do ms9po400
' bedzie musialo byc to polaczone zdecydowanie z implementacja ekranu ms9pop00
' co ciekawe juz to implementacja jest na miejscu
' i wiele faktycznie nie trzeba
' tylko flaga na bardziej abstrakcyjnym obiekcie
' dzieki ktoremu przechowamy odpowiednie info

'
' blad w formie wejsciowym po kliknieciu more
' i rozwinieciu listy mamy guziki >> i <<
' okazuje sie ze usuwaja z listy doslownie wszystko nawet te ktore powinny zostac jako nieusuwalne :(
' nalezalo rowniez zrobic fix na tej wersji

'4.02.10
' -------------------------------------------------------------
' teraz kolejna sprawa zwiazana z konfiguracjami komputerow
' pojawia sie srogi temat mismatchy na innych komputerach niz moich
' widac bez wszystkich jawnych dim'ow sie nie obedzie
'
' referencja na dodatkowych bibliotek z chartem i takie tam
' zostawilem sobie kiedys na innych projektach no i przecko inni ludzi nie posiadaja takowych ustawien

'4.02.09
' -------------------------------------------------------------
' wersja poprawiona z not yet received poniewaz sciaga dane z ms9po400 i zk7ppum0
' jednak ze zrobilem logike nie uwzgledniania duplikatow tylko na regularach
' not yet received sie duplikawolo
' od wersji 4.02.09 ten problem zostal usuniety
'
'


'4.02.08
' -------------------------------------------------------------
' no i niestety musialem zmienic implementacje na nowo zeby znow patrzyla sobie
' i na zk7ppus0 jak i na ms9po400 poniewaz moze sie okazac ze pomimo tego ze czesc jest pusowa
' i tak koniec koncow moze ktos zrobic help ship a przeciez to jest asn :(
' glowny zmiany w klasie DailyIteration w sub:
' fillTransitCollectionFromMs9po400 gdzie doszedl drugi argument opcjonalny
' kiedy to ms9po400 dziala samodzielnie i kiedy dziala razem z zk7ppus0
' i robi match na juz istniejacych danych


'4.02.07
' -------------------------------------------------------------
' first runout fixed (assumption)
' autofit on first runout
' still static args on formula first runout look on calendar 7th of April
' ' -
' ' -
' wlasciwie dziala bez zarzutu juz mozna pomalu sie zajac tematem run from break

'4.02.06
' -------------------------------------------------------------
' change on first runout & test run on change values whree first runout how now one arg
' cos jest nie tak!
' trza naprawic first runout robi blad typu #Value!


'4.02.05
' -------------------------------------------------------------
' change on first runout

'4.02.04
' -------------------------------------------------------------
' zmiany na definicjach kolorow raportu wyjsciowego
' koloru o ref yellow, grey staly sie odniesieniem (kolumna J)
' kolory faktycznie sterujace wygladem sa od teraz w kolumnie M
' komorki glowne:
' od M1 do M5:
' - primaryColor
' - secondaryColor
' - minusColor
' - warningColor
'
'
' dodatkowo zmiana jest guzik w formularzu More/Less odpowiedzialny za odkrywanie dodatkowych funkcji wstepnej konfiguracji ffl
' chodzi przede wszystkim o wyswietlanie takich danych jak weeknum
' zmienianie danych common poprzez dwa comboboxy
' limitowanie historii pusow
' jak daleko chcemy patrzec w przeszlosc ekranu zk7ppus0
'
'
' dodatkowa implementacja ekranu MS3P9800
' sama klasa w sobie jest uboga jednak jest to glowny warunek spradzajacy czy w ogole mamy przejrzec dane na ekranie ms9po400

'4.02.03
' -------------------------------------------------------------
' fix na klasie dynamic kolors w przyapdku gdy chcemy zobaczyc tylko jedno linie na raporcie

'4.02.02
' -------------------------------------------------------------
' dodana klasa dynamic kolors ktora przelicza ruchliwe kolory


'4.02.01
' -------------------------------------------------------------
' kolejna podwersja podczas pracy implementacyjnej
' z racji tego ze skonczyly mi sie liczby 4.01.x
' musze przejsc kawalek dalej z numeracja, ale tym razem bede chytry i trzecia instancje versjonowania dam dwucyfrowa
' zaczynajac prace z ta wersja jestem gotowy do przygowoania implementacji kolorow dynamicznych
' jesli chodzi o implementacje pusow to w dalszym ciagu jest otwarta kwestia
' jak mam podjesc do ekranu ms9ph100, czyli ekranu historii
' aktualnie jest zrobione tak, ze:
' - recv na zero jest ignorowany
' - recv na dokladna wartosc kasuje pus
' - recv na wartosc inna kasuje pus
' - asn dr kasuje pus


' 4.01.9
' wersja na ktorej zaczynam budowac implementacje pod pobieranie pusow
' chodzi o to ze tym razem olewam temat zwiazany z wchodzeniem za kazdym razem na ppus0 i ms9po400 za kazdym razem
' duplikowalem dane a potem wyrzucalem ich nadmiar
' zatrzymalem sie na szablonie wyrzucania danych na podstawie tego co widzimy na ekranie historii
' usunlaem z pickupHandler kolekcje z ph100 ekranu historii
' tj. tam jest porownywanie bezposrednie


' od tej pory ostatni raport czy tez ostatnia odslona bedzie od gory
' 4.01.8
' =============================================================================
' walka z kolekcja ic ktora przechowuje kolejne ajajtemy
' problem w tym ze metoda addcolorflavour ktora z reszta nie do konca swoja nazwa przekazuje co wlasciwie robi a robi o wiele wiecej
' trzeba bedzie sie zastanowic nad sensem jej istnienia i ewentualnego przemeblowania jej implementacji na poziom zdecydowanie bardziej
' oop a nie takie fiu bziu niesamowite cos co na dluzsza mete nie ma sensu
' zatem:
' kolekcja ic jest polem obiektow ktore dziedzicza interfejs typu IIteration
' jednak sam obiekt instancji takowej pracuje w metodzie repRun klasy IReport
' jest tam kolejno uruchomiony:
' create full iteration
' add color flavour
' problem pojawia sie miedzy nimi poniewaz
' dalbym sobie glowe uciac ze bedzie ta pierwsza wsadza dane do kolekcji ic
' natomiast od tej drugiej juz te dane powinny byc widoczne a nie sa
' co jest super dziwne i dlatego na dzien 2014-11-14 add color flavour traktuje ic jako kolekcje pusta
'
'
'
' co zostalo zrobione
' narazie pusy z po400 czyli wlasciwie logika tylko pod po400
' dodane kolory
' ograniczenie czarnego
' dodatnie weeknum nad data konkrentego itemu
' naprawienie pustych dat, jesli w zk7pdrqm zaczynaja sie dane duzo pozniej
' fillTransitCollectionByPuses
' to procedura bedzie posiadala logike dzialania w oddelegowanym obiekcie typu PickupHandler - nowo powstaly obiekt
' ktory jeszcze nie mial miejsca w dotychczasowej implementacji ff'a
'
' kolejna nowa klasa jest Komentarz - do niego beda delegowane wszelkie zadania zwiazane z tworzeniem komentarza w raporcie
' =============================================================================



' 4.01
' pisze wlasciwie od poczatku - dalej implementacja oparta bedzie na OOP
' jednak duzo bardziej przemyslnie tak, aby nie trzeba bylo zbyt duzo pisac
'
' no i agree
' nadmiar kodu w wersjach trzeciej generacji wbijal mnie w ziemie
' =============================================================================

' 4.01.3
' przesuniecie metod prywatnych z DailyReport na DailyIteration jako przygotowanie rusztowania
' przesuniecie wynika z intuicyjnego podjescia do sprawy, gdzie strategia sprawdzenia plt na poczatku wydawala sie zgodna z wartstwa raportu
' jednak w miare rozwoju projekty czesciowy fill iteracji przesuniety zostal jako metoda wewnetrzna co wymoglo na checkach na plantach aby rowniez staly
' sie metodami wewnetrznymi itemow iteracji aby metoda czesciowego fillu miala swobodny dostep bez nadmiernych akrobacji
' =============================================================================
' 4.01.4
' kontynuacja rozwoju metod uzupelnniajacych transit
' podjeta decyzja zmiany downloadu:
' tj.
' osobno sciagam dane z zk7ppus0
' albo z ms9po400 tylko i wylaczenie wtedy gdy mamy do czynienia z asnami
' a  tak staramy sie trzymac tylko pusowych ekranow
' =============================================================================

' 4.01.5
' pierwsze kroki z layoutem
' dobieranie koloru jak i koncepcji ustawienia danych - formularz wejsciowy co chce widziec na raporcie z popa?
' dopisanie brakujacych popowych parametrow
' =============================================================================

' 4.01.6
' wersja pracy nad layoutem nowego ff lighta :D
' =============================================================================
' 4.01.7
' zatrzymalem sie na tym ze kolekcja ic w klasie FireFlakeLayoutDaily
' nie posiada zadnego contentu
' - tak jakby byla pusta kolekcja nie wiadomo dlaczego
' =============================================================================




