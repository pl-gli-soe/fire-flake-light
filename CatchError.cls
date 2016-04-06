VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CatchError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function dataAvail(screen_name As String, infocode As String) As Boolean
    
    
    
    dataAvail = True
    If screen_name = "ms9pop00" Then
        ' I4266: DATA NOT FOUND FOR SCAN REQUEST - for example
        If infocode = "I4266" Then
            dataAvail = False
        End If

    ElseIf screen_name = "zk7pdrqm" Or screen_name = "zk7pwrqm" Then
    
        ' R6116: PART NUMBER NOT FOUND
        ' R6105: MUST INQUIRE BEFORE PF11 CAN BE USED
        If infocode = "R6116" Or infocode = "R6105" Then
            dataAvail = False
        End If
    End If
End Function


Public Function catchIfInDailyRqmsThereAreNoDates(strFromMgo As String) As Boolean
    If strFromMgo = "" Then
        catchIfInDailyRqmsThereAreNoDates = True
    Else
        catchIfInDailyRqmsThereAreNoDates = False
    End If
End Function

Public Function catchIfPartNotFoundInHistoryScreen(strFromMgo As String) As Boolean
    catchIfPartNotFoundInHistoryScreen = False
    
    
    If strFromMgo <> "I4017" Then
        catchIfPartNotFoundInHistoryScreen = True
    Else
        catchIfPartNotFoundInHistoryScreen = False
    End If
End Function
