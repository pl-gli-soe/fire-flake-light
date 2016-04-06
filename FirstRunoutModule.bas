Attribute VB_Name = "FirstRunoutModule"
' to jest jedna z najwazniejszych funkcji calego projektu
' zatem totez dlatego ma nawet swoj wlasny modul
' co by nie bylo watpliwosci
' gdybym chcial powaznie
' rozmyslac o ewentualnej przyszlej rozbudowie tej oto funkcji
Public Function firstRunout(r As Range) As String
    firstRunout = ""
    
    'wiersz_w_ktorym_znajduja_sie_daty = 3
    'wwkzsd = wiersz_w_ktorym_znajduja_sie_daty
    
    Dim sh As Worksheet, rng As Range, ebal_flag As Range
    Set sh = r.Parent
    Set rng = sh.Range("a3").End(xlToRight)
    Set ebal_flag = rng.Offset(1, 3)
    ' Debug.Print ebal_flag
    
    Do
    
        If sh.Cells(r.Row, ebal_flag.Column) < 0 Then
            firstRunout = ebal_flag.Offset(-1, -2)
            Exit Function
        End If
        Set ebal_flag = ebal_flag.Offset(0, 3)
    Loop Until Trim(ebal_flag) = ""
    
    
    ' to jest przeklamanie!
    ' firstRunout = ebal_flag.Offset(-1, -5)
    
End Function
