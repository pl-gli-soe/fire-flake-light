VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IItem


Private mdate As Date
Private rqm As Long
' tylko jako handler
Private t As ITransit
Private transCollection As Collection
Private ebal As Range

' te pole powstalo w zwiazku z potrzeba na koncowych akcjach tworzenia layoutu
' wpisywanie danych w komorki fizycznie zajmowalo zbyt duzo czasu
' zatem dobrze by bylo aby istniala juz gotowa wartosc z ktorej po prostu zbieramy dane
Private sumOfQty As Long



Private Sub Class_Initialize()
    Set transCollection = New Collection
    sumOfQty = 0
End Sub

Private Sub Class_Terminate()
    Set transCollection = Nothing
End Sub

Private Sub IItem_addRqm(v As Long)
    rqm = rqm + CLng(v)
End Sub

Private Sub IItem_addTtransit(pus As ITransit)
    transCollection.Add pus
    sumOfQty = sumOfQty + CLng(pus.mQty)
End Sub

Private Function IItem_getDate() As Date
    IItem_getDate = CDate(mdate)
End Function

Private Function IItem_getEBAL() As Range
    Set IItem_getEBAL = ebal
End Function

Private Function IItem_getInTransits() As Collection
    Set IItem_getInTransits = transCollection
End Function

Private Function IItem_getRqm() As Long
    IItem_getRqm = CLng(rqm)
End Function

Private Function IItem_getSumOfQty() As Long
    IItem_getSumOfQty = CLng(sumOfQty)
End Function

Private Function IItem_getTransitItem() As ITransit
    Set IItem_getTransitItem = t
End Function

Private Sub IItem_setDate(d As Date)
    mdate = CDate(d)
End Sub


Private Sub IItem_setEBAL(r As Range)
    Set ebal = r
End Sub

Private Sub IItem_setRqm(v As Long)
    rqm = CLng(v)
End Sub
