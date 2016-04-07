VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DownloadingRulesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private beg_of_list As Range
Private end_of_list As Range
Private zasieg_from_beg_to_end As Range
Private rqm_source As Range
Private where_am_i As Integer
Private r As Range


' !!!!!!!!!
' to byla globalna zmienna zobaczymy jak sobie radzi w srodowisku jako komponent
' ------------------------------------------------------------------------------------
Public DEOC As DynamicEventOnControls
' ------------------------------------------------------------------------------------


Public ile_bedzie_plantow As Long
Public overall_offset As Long

Public c As Control
Public tbc As Control
Public tb As TextBox
Public cb As ComboBox
Public lbl As Label

Public control_events_coll As Collection


Public Sub okresl_wielkosc_forma_i_przesun_guziki_w_dol()
    
    overall_offset = CLng(OFFSET_FOR_NEW_PLT) * CLng(ile_bedzie_plantow)
    
    ' tutaj wielkosc forma
    ustaw_wielkosc_calego_forma
    ' tutaj guziki ida na dol
    guziki_na_dol
End Sub

Private Sub guziki_na_dol()
    DownloadingRulesForm.BtnAdd.Top = CLng(INIT_RULES_FORM_BTN_TOP) + CLng(overall_offset)
    DownloadingRulesForm.BtnRemove.Top = CLng(INIT_RULES_FORM_BTN_TOP) + CLng(overall_offset)
    DownloadingRulesForm.BTNHide.Top = CLng(INIT_RULES_FORM_BTN_TOP) + CLng(overall_offset)
End Sub

Private Sub ustaw_wielkosc_calego_forma()
    DownloadingRulesForm.Width = INIT_RULES_FORM_WIDTH
    DownloadingRulesForm.Height = INIT_RULES_FORM_HEIGHT + CLng(overall_offset)
    'DownloadingRulesForm.Frame.Width = INIT_RULES_FORM_FRAME_WIDTH
    'DownloadingRulesForm.Frame.Height = INIT_RULES_FORM_FRAME_HEIGHT + CLng(overall_offset)
End Sub


Public Sub iteracja(x As Integer)


    Set DEOC = Nothing
    Set DEOC = New DynamicEventOnControls

    Set tbc = DownloadingRulesForm.Controls.Add("Forms.TextBox.1", "Name" & CStr(x), True)
    ' DownloadingRulesForm.Controls.Add "Forms.TextBox.1", "Name1", True
    tbc.Left = TEXTBOX_PLT_0_LEFT
    tbc.Top = TEXTBOX_PLT_0_TOP + (OFFSET_FOR_NEW_PLT * CLng(x))
    tbc.Width = TEXTBOX_PLT_0_W
    tbc.Height = TEXTBOX_PLT_0_H
    
    ' Set tb = c
    tbc.Text = getPlants().Item(Int(x))
    
    Set DEOC.textBoxControl = tbc

    Set c = DownloadingRulesForm.Controls.Add("Forms.ComboBox.1", "Kombo" & CStr(x), True)
    c.Left = TEXTBOX_RQM_0_LEFT
    c.Top = TEXTBOX_RQM_0_TOP + (OFFSET_FOR_NEW_PLT * CLng(x))
    c.Width = TEXTBOX_RQM_0_W
    c.Height = TEXTBOX_RQM_0_H
    
    Set cb = c
    For Each r In getRqmSource
        cb.AddItem CStr(r)
    Next r
    
    
    cb.Value = getPlants().Item(Int(x)).Offset(0, 1)
    
    Set DEOC.Control = cb
    DEOC.setHandlerName CStr(x)
    
    control_events_coll.Add DEOC
    whereAmI = Int(x)
    
    Set c = DownloadingRulesForm.Controls.Add("Forms.Label.1", "Label" & CStr(x), True)
    c.Left = LABEL_CMNT_0_L
    c.Top = LABEL_CMNT_0_T + (OFFSET_FOR_NEW_PLT * CLng(x))
    c.Width = LABEL_CMNT_0_W
    c.Height = LABEL_CMNT_0_H
    
    ' Set lbl = c
    ' r juz wczesniej zostalo zdimowane :)cda
    For Each r In getRqmSource
        ' cb.AddItem CStr(r)
        If getPlants().Item(Int(x)).Offset(0, 1) = r Then c.Caption = r.Offset(0, 1)
    Next r
    
End Sub


Public Property Get whereAmI() As Integer
    whereAmI = where_am_i
End Property

Public Property Let whereAmI(arg As Integer)
    where_am_i = arg
End Property


Public Function getPlants() As Range
    Set getPlants = zasieg_from_beg_to_end
End Function

Public Function getRqmSource() As Range
    Set getRqmSource = rqm_source
End Function

Private Sub Class_Initialize()
    inicjacja
End Sub

Public Sub inicjacja()

    Set beg_of_list = ThisWorkbook.Sheets("register").Range("r3")
    
    If ThisWorkbook.Sheets("register").Range("r4") = "" Then
        Set end_if_list = ThisWorkbook.Sheets("register").Range("r3")
    Else
        Set end_if_list = beg_of_list.End(xlDown)
    End If
    
    Set zasieg_from_beg_to_end = ThisWorkbook.Sheets("register").Range(beg_of_list, end_if_list)
    
    
    
    Set rqm_source = ThisWorkbook.Sheets("register").Range("u2")
    Set rqm_source = ThisWorkbook.Sheets("register").Range(rqm_source, rqm_source.End(xlDown))
    
    
    Set control_events_coll = New Collection
End Sub

Private Sub Class_Terminate()
    Set control_events_coll = Nothing
End Sub
