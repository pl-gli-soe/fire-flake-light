Attribute VB_Name = "VersionModule"
' Version Module

' no to start
' modul ten powstal od czwartej generacji ffa
' w mysl o tym czego wlasciwie potrzebuje uzytkownik musze pomyslec ile faktycznie jest narzedzi potrzebnych a ile nie
' odchudzenie ffa rowniez wchodzi w gre aby chodzil odrobine szybciej

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
