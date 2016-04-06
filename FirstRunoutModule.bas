Attribute VB_Name = "FirstRunoutModule"
' to jest jedna z najwazniejszych funkcji calego projektu
' zatem totez dlatego ma nawet swoj wlasny modul
' co by nie bylo watpliwosci
' gdybym chcial powaznie
' rozmyslac o ewentualnej przyszlej rozbudowie tej oto funkcji
Public Function firstRunout(r As Range) As Date
    firstRunout = CDate(Now)
    
    'wiersz_w_ktorym_znajduja_sie_daty = 3
    'wwkzsd = wiersz_w_ktorym_znajduja_sie_daty
    
    Dim sh As Worksheet, rng As Range, ebal_flag As Range
    Set sh = r.Parent
    Set rng = sh.Range("a3").End(xlToRight)
    Set ebal_flag = rng.Offset(1, 3)
    Debug.Print ebal_flag
    
    Do
    Loop Until ebal_flag
    
End Function
