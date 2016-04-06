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
