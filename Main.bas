Attribute VB_Name = "Main"
' jest to glowna metoda ktora bedzie stand alone w przypadku podpinania do guzika tj.
' bedzie wpisana jako jedyna do subroutine ktore bedzie podpiete do guzika juz bezposrednio bez zbednych dodatkowych zabiegow
' jej wszsytkie argumenty maja w pelni ogarnac konfiguracje run ff light
' myslalem nad kombinacja alpejska w stylu zeby user mial mozliwosc konfigurowalnosci widoku ale chyba bylo by to przedobrzone :D
' jeszcze obaczym - napisal to ja 2014 wrzesien 22.
Public Sub runReport(t As RUN_TYPE, p_limit As Date, l As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Date)


    Application.ScreenUpdating = False

    Dim ffl As FireFlakeLight
    Set ffl = New FireFlakeLight
    Set sh = New StatusHandler


    If st = FROM_THE_BEGINNING Then
        If t = DAILY Then
            ffl.runDaily CDate(p_limit), l, st, CDate(daily_rqm_limit)
        ElseIf t = HOURLY Then
            ffl.runHourly p_limit, l, st, CDate(daily_rqm_limit)
        ElseIf t = WEEKLY Then
            ffl.runWeekly p_limit, l, st, CDate(daily_rqm_limit)
        End If
    ElseIf st = CONTINUE_BROKEN_ONE Then
    
        ' ten tutaj jest bystry na tyle zeby sam siebie skonfigurowac i pociagnac temat samemu
        ' :)
        ffl.continueBrokenReport
    End If
    
    
    Set ffl = Nothing
    Set sh = Nothing
    Application.ScreenUpdating = True
End Sub




Public Sub run_ff(ictrl As IRibbonControl)
    MainForm.show
End Sub

Public Sub reset_report_inner()
    ' tu ma byc reset jako taki dla odswiezenia dynamicznych kolorow
    ' teraz kwestia tylko z jakim rodzajem raportu mamy do czynienia
    Dim ash As Worksheet, dc As IDynamicColors
    If ThisWorkbook.FullName = ActiveWorkbook.FullName Then
        Set ash = ActiveSheet
        
        If CStr(ash.Range("b4")) = "Part #" And CStr(ash.Range("c4")) = "Plant" Then
            Set dc = New DailyDynamicColors
            dc.assignDynamicColorsrange
            dc.recalcColors
        End If
        
    Else
        Set ash = Nothing
    End If
End Sub

Public Sub reset_report(ictrl As IRibbonControl)
    reset_report_inner
End Sub
