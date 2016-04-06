Attribute VB_Name = "TestModule"
Public Sub showStatusForm()
    Dim sh As StatusHandler
    Set sh = New StatusHandler
    
    sh.init_statusbar 100
    
    sh.show
    
    For x = 1 To 100
        sh.progress_increase
        Sleep 100
    Next x
    
    
    sh.hide
    
    Set sh = Nothing
End Sub


Public Sub runDaily()
    Dim ffl As FireFlakeLight
    Set ffl = New FireFlakeLight
    
    ffl.runDaily Now + 100, LIST_LAYOUT, FROM_THE_BEGINNING, Now + 100
End Sub



Public Sub test_autofit()

    Dim rng As Range, sh As Worksheet

    Set sh = ThisWorkbook.ActiveSheet
    Set rng = sh.Range("a3").End(xlToRight)
    
    Set rng = sh.Range(sh.Cells(4, 2), sh.Cells(4, rng.Column - 1))
    rng.EntireColumn.AutoFit
End Sub

Private Sub filter_on()
'
' filter_on Macro
'

'
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("B4").Select
End Sub

