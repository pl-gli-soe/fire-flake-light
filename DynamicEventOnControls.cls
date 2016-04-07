VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicEventOnControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents handler As MSForms.ComboBox
Attribute handler.VB_VarHelpID = -1
Private WithEvents nameh As MSForms.TextBox
Attribute nameh.VB_VarHelpID = -1
Private pEventSelected As Control
Private handler_name As Integer



Private Sub handler_Change()
    ' Debug.Print "tutaj arkusz powinien sie dopasowac"
    
    ThisWorkbook.Sheets("register").Cells(Int(handler_name) + 2, CONFIG_REG_PLT_COLUMN + 1) = CStr(handler.Value)
    
    
    ' prawie jak konstruktor - tyle ze nie kasuje danych nieodswiezanych :)
    ' co smieszne implementacja wyglada tak samo jak w new DRH
    'drh.inicjacja
    'drh.okresl_wielkosc_forma_i_przesun_guziki_w_dol
    'DownloadingRulesForm.Repaint
End Sub


Public Sub setHandlerName(s As Integer)
    handler_name = Int(s)
End Sub

Public Function getHandlerName() As String
    getHandlerName = handler_name
End Function


Public Property Set Control(p As MSForms.ComboBox)
    Set handler = p
End Property

Public Property Set textBoxControl(p As MSForms.TextBox)
    Set nameh = p
End Property



Private Sub nameh_Change()
    ThisWorkbook.Sheets("register").Cells(Int(handler_name) + 2, CONFIG_REG_PLT_COLUMN) = CStr(nameh.Value)
End Sub
