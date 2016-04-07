Attribute VB_Name = "TestModule"
Public Sub cebeer_test()
    Dim c As ContinueBrokenReportHandler
    Set c = New ContinueBrokenReportHandler
    
    c.setPusLimit Range("a1")
End Sub


Public Sub continueReport()
    Dim ffl As FireFlakeLight
    Set ffl = New FireFlakeLight
    ffl.continueBrokenReport LIST_LAYOUT, CONTINUE_BROKEN_ONE
    
End Sub


Private Sub closingLine()
'
' closingLine Macro
'

'
    Range("Q5:Q10").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -3618616
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -3618616
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -3618616
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = -3618616
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


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

Private Sub freezePanesTest()
'
' freezePanesTest Macro
'

'
    Range("D5").Select
    ActiveWindow.freezePanes = True
End Sub

Private Sub shrinkColumns()
'
' shrinkColumns Macro
'

'
    Columns("B:N").Select
    Selection.ColumnWidth = 7.43
End Sub

