VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Main Form"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClose_Click()
    Application.ThisWorkbook.Close
End Sub

Private Sub BtnHide_Click()
    Me.hide
End Sub

Private Sub BtnMoreLess_Click()
    If Me.BtnMoreLess.Caption Like "*More*" Then
        Me.Height = 440
        Me.BtnMoreLess.Caption = "Less"
    Else
        Me.Height = 122
        Me.BtnMoreLess.Caption = "More"
    End If
        
End Sub

Private Sub BtnMoveAllToLeft_Click()

    fillAllPopsDataByChar "x"
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init

End Sub

Private Sub BtnMoveAllToRight_Click()


    fillAllPopsDataByChar ""
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub fillAllPopsDataByChar(ch As String)
    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
    
    Do
    
        r.Offset(0, 1) = ch
        Set r = r.Offset(1, 0)
    Loop While r <> ""
End Sub

Private Sub btnMoveToLeft_Click()
    change_register_workhseet "x"
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub BtnMoveToRight_Click()
    change_register_workhseet ""
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub BtnRunDaily_Click()
    Me.hide
    
    Dim wybor_typu_layoutu As LAYOUT_TYPE
    
    If Me.OptionButtonList.Value Then
    
        wybor_typu_layoutu = LIST_LAYOUT
    ElseIf Me.OptionButtonCoverage.Value Then
    
        wybor_typu_layoutu = COV_LAYOUT
    ElseIf Me.OptionButtonBox.Value Then
    
        wybor_typu_layoutu = BOX_LAYOUT
    End If
    
    If Me.DTPickerPUSLimit.Enabled Then
        ThisWorkbook.Sheets("register").Range("pusLimit") = CDate(Me.DTPickerPUSLimit.Value)
    Else
        ThisWorkbook.Sheets("register").Range("pusLimit") = CDate(Me.DTPickerPUSLimit.Value) + 100
    End If
    
    If Me.DTPickerRQMLimit.Enabled Then
        ThisWorkbook.Sheets("register").Range("rqmLimit") = CDate(Me.DTPickerRQMLimit.Value)
    Else
        ThisWorkbook.Sheets("register").Range("rqmLimit") = CDate(Me.DTPickerRQMLimit.Value) + 100
    End If
    
    
    ThisWorkbook.Sheets("register").Range("LAYOUT_TYPE") = wybor_typu_layoutu
    ThisWorkbook.Sheets("register").Range("RUN_TYPE") = DAILY
    ThisWorkbook.Sheets("register").Range("START_TYPE") = FROM_THE_BEGINNING
    
    runReport DAILY, CDate(ThisWorkbook.Sheets("register").Range("pusLimit")), wybor_typu_layoutu, FROM_THE_BEGINNING, CDate(ThisWorkbook.Sheets("register").Range("rqmLimit"))
End Sub

Private Sub BtnRunWeekly_Click()
    Me.hide
    MsgBox "not yet implemented"
    Me.show
End Sub

Private Sub CheckBoxPUSLimit_Click()
    If Not Me.CheckBoxPUSLimit.Value Then
        Me.DTPickerPUSLimit.Enabled = False
    Else
        Me.DTPickerPUSLimit.Enabled = True
    End If
End Sub

Private Sub CheckBoxRQMLimit_Click()

    If Not Me.CheckBoxRQMLimit.Value Then
        Me.DTPickerRQMLimit.Enabled = False
    Else
        Me.DTPickerRQMLimit.Enabled = True
    End If
End Sub

Private Sub CheckBoxWeekNum_Click()
    With Me.CheckBoxWeekNum
        If .Value = True Then
            ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 1
        Else
            ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 0
        End If
    End With
End Sub

Private Sub CommandButton1_Click()
    Me.hide
    MsgBox "not yet implemented"
    Me.show
End Sub

Private Sub change_register_workhseet(s As String)

    Dim r As Range

    If s = "" Then
        
        For x = 0 To Me.ListBoxInCellLeft.ListCount - 1
            If Me.ListBoxInCellLeft.Selected(x) Then
                tmp = Me.ListBoxInCellLeft.List(x)
                
                
                
                Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
                
                Do
                    If CLng(r.Interior.Color) <> CLng(ThisWorkbook.Sheets("register").Range("black")) Then ' as black
                        If CStr(tmp) = CStr(r) Then
                            r.Offset(0, 1) = s
                        End If
                    End If
                    Set r = r.Offset(1, 0)
                Loop While r <> ""
                
            End If
        Next x
    ElseIf s = "x" Then
    
    
        For x = 0 To Me.ListBoxInCommentRight.ListCount - 1
            If Me.ListBoxInCommentRight.Selected(x) Then
                tmp = Me.ListBoxInCommentRight.List(x)
                
                
                
                Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
                
                Do
                    If CLng(r.Interior.Color) <> CLng(ThisWorkbook.Sheets("register").Range("black")) Then ' as black
                        If CStr(tmp) = CStr(r) Then
                            r.Offset(0, 1) = s
                        End If
                    End If
                    Set r = r.Offset(1, 0)
                Loop While r <> ""
                
            End If
        Next x
    End If
End Sub




Private Sub UserForm_Initialize()


    ' dates now
    Me.DTPickerPUSLimit = Now
    Me.DTPickerRQMLimit = Now
    
    Me.Height = 122



    ' week #
    Me.CheckBoxWeekNum.Value = True
    ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 1


    ' layout type
    ' ============================================
    
    With Me.LayoutTypeFrame
        .Enabled = False
    End With
    
    Me.OptionButtonList.Value = True
    Me.OptionButtonCoverage.Value = False
    Me.OptionButtonBox.Value = False
    
    ' ============================================


    ' history limit
    ' ============================================
    Me.ComboBoxHistoryLimit.Clear
    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("BegOfHistoryLimitRange")
    
    Do
        Me.ComboBoxHistoryLimit.AddItem r
        Set r = r.Offset(1, 0)
    Loop While r <> ""
    
    Me.ComboBoxHistoryLimit = assign_default_value()
    ' ============================================
    


    ' limitacje
    Me.DTPickerPUSLimit.Enabled = False
    Me.DTPickerRQMLimit.Enabled = False
    Me.CheckBoxPUSLimit.Value = False
    Me.CheckBoxRQMLimit.Value = False
    
    
    
    ' tutaj zabawa z konfiguracja danych z popa
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
    
End Sub

Private Function assign_default_value()

    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("BegOfHistoryLimitRange")
    
    Do
        If r.Offset(0, -1).Value = "default" Then
            assign_default_value = r
            Exit Function
        End If
        Set r = r.Offset(1, 0)
    Loop While r <> ""

End Function


Private Sub set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init()

    Me.ListBoxInCellLeft.Clear
    Me.ListBoxInCommentRight.Clear

    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
    
    Do
    
        If r.Offset(0, 1) = "x" Then
            Me.ListBoxInCellLeft.AddItem r
        Else
            Me.ListBoxInCommentRight.AddItem r
        End If
        Set r = r.Offset(1, 0)
    Loop While r <> ""
End Sub

