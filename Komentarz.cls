VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Komentarz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public typeOfComment As COMMENT_TYPE


Public Sub assignTransitsIntoComment(r As Range, c As Collection)

    If c.COUNT > 0 Then
        typeOfComment = IN_TRANSIT
        
        r.ClearComments
        txt = CStr(prepareComment(c))
        If Trim(txt) <> "" Then
            r.AddComment CStr(txt)
            ' wartosci dobrane heurstycznie
            r.comment.Shape.Width = 100
            r.comment.Shape.Height = 120 * (c.COUNT)
        End If
        
        
    End If
End Sub

Private Function prepareComment(c As Collection) As String

    ' c to kolekcja kolejnych tranzytow
    txt = ""
    Dim t As ITransit
    For Each t In c
        If t.mNotYetReceived Then
            txt = txt & _
            "eda: " & CStr(t.mDeliveryDate) & Chr(10) & _
            "eta: " & CStr(t.mDeliveryTime) & Chr(10) & _
            "sdate: " & CStr(t.mPickupDate) & Chr(10) & _
            "ptime: " & CStr(t.mPickupTime) & Chr(10) & _
            "name: " & CStr(t.mName) & Chr(10) & _
            "qty: " & CStr(t.mQty) & Chr(10) & _
            "route: " & CStr(t.mRoute) & Chr(10) & _
            "status: " & CStr(t.mST) & Chr(10) & _
            "trlr: " & CStr(t.mTRLR) & Chr(10) & _
            "--------------------" & Chr(10)
        End If
    Next t
    
    prepareComment = txt

End Function
