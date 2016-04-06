Attribute VB_Name = "GlobalModule"
#If VBA7 Then
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As LongPtr
#Else
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
#End If



' delay time
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' pobierz dynamiczna library
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)


Global Const LETTERS = 26
Global Const MAX_COLUMNS = 16384 ' ostatnia kolumna
Global Const C_HOUR = (0.041 + 0.001 * (2 / 3))
Global Const INITIAL_TIMING_FOR_ONE_PN = 6

Global sh As StatusHandler

Public Enum ENUM_LEFT_RIGHT_LISTBOX
    MOVE_TO_LEFT_LISTBOX
    MOVE_TO_RGHT_LISTBOX
End Enum


Public Enum RUN_TYPE
    DAILY
    WEEKLY
    HOURLY
End Enum

Public Enum LAYOUT_TYPE
    LIST_LAYOUT
    COV_LAYOUT
    BOX_LAYOUT
End Enum

Public Enum START_TYPE
    FROM_THE_BEGINNING
    CONTINUE_BROKEN_ONE
End Enum


Public Enum ITERATION_CONFIG
    CONFIG_ONE
    CONFIG_TWO
    CONFIG_THREE
End Enum

Public Enum COMMENT_TYPE
    IN_TRANSIT
    DATA_FROM_POP
End Enum


Public Function MGO_active(m As MGO) As Boolean


    MGO_active = False
    
    If m Is Nothing Then
        MGO_active = False
        MsgBox "mgo session is nothing!"
        Exit Function
    End If
    
    If m.actualScreen <> "" Then
        MGO_active = True
    End If
End Function


Public Function chrx(col As Integer, Optional ByRef s As Box) As String

    If col <= MAX_COLUMNS And col > 0 Then
    If s Is Nothing Then
        Set s = New Box
    End If
    
    If col > LETTERS Then
        s.counter = s.counter + 1
        If s.counter = 26 Then
        ' wersja prostsza
            s.counter = 0
            s.scope = s.scope + 1
        End If
        chrx = chrx(col - LETTERS, s)
    Else
        If s.counter = 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + s.counter) + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope <> 0 Then
            chrx = chrx + Chr(64 + s.scope) + Chr(64 + s.counter) + Chr(64 + col)
        End If
    End If
    Else
        MsgBox "out of scope mf! MAX_COLUMNS = 16384"
    End If
    
   
End Function


Public Sub refresh_register_worksheet()

End Sub



