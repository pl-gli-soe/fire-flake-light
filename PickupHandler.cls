VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickupHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' obydwie kolekcje sa kontenerami typu ITransit
Private ppum0coll As Collection
Private t As ITransit

Private HOW_MANY_DAYS_FOR_PPUS0 As Range


Public Sub takeDataFromPPUM0(m As MGO)

    
    Do
        For x = 0 To 19 - 7
        
            If Trim(m.pZK7PPUM0.plt(Int(x))) <> "" Then
            
                Set t = Nothing
                Set t = New Transit
                
                With m.pZK7PPUM0
                    t.mDeliveryDate = CDate(m.convertToDateFromZK7PPUM0Date(.DEL_DATE(Int(x))))
                    t.mDeliveryTime = CDate(Mid(.DEL_WINDW(Int(x)), 6, 2) & ":" & Right(.DEL_WINDW(Int(x)), 2))
                    t.mOnPo400 = False
                    t.mName = CStr(.PUS_NBR(Int(x)))
                    t.mNotYetReceived = True
                    t.mQty = CLng(.PART_QTY(Int(x)))
                    t.mPickupDate = CDate(m.convertToDateFromZK7PPUM0Date(.PU_DATE(Int(x))))
                    t.mRoute = CStr(.route(Int(x)))
                    
                    ppum0coll.Add t
                    Set t = Nothing
                End With
            End If
        Next x
        
        If (m.getCode("zk7ppum0") = CStr("R6086")) Or (m.getCode("zk7ppum0") = CStr("R6123")) Then
            Exit Do
        End If
        
        m.sendKeys "<pf8>"
    Loop While True
End Sub


Public Sub takeDataFromPH100(m As MGO, mcd As CommonData)

    m.sendKeys "<Clear>"
    m.sendKeys "ms9ph100 <Enter>"
    With m.pMS9PH100
        .plt = mcd.plt
        .pn = mcd.pn
        .setDTYP "RECV", "ASN"
        .DATE_FROM = m.convertToMgoShortDate(CDate(Now - HOW_MANY_DAYS_FOR_PPUS0))
        m.submit
        
        
        ' teraz bardzo wazna rzecz
        ' zeby to dzialalo poprawnie nalezy sprawdzic czy w ogole czesc jest znaleziona przez ekran historii
        
        ' check if there are no  issues like : I4017: PART NOT FOUND
        
        Dim ce As CatchError
        Set ce = New CatchError
        
        If ce.catchIfPartNotFoundInHistoryScreen(.infocode) Then
            Do
                For x = 0 To 9
                    
                    ' kolejne wiersze z ph100 - ekran historii
                    If ppum0coll.COUNT > 0 Then
                        For Each t In ppum0coll
                            
                            
                            
                            ' 3 warunki sprawdzajace czy recv jest na konkretny sid wraz z odowiednia wartoscia
                            ' TERAZ SUPER WAZNE - ZACHOWANIE LOGIKI DLA RECV JAK I DLA ASN JEST IDENTYCZNE WIEC
                            ' JESLI NIE ZAMIERZAM W POZNIEJSZYM CZASIE ZMIENIAC CZEGO KOLWIEK
                            ' ALBO JESLI CZEGOS REWOLUCYJNEGO SIE NIE DOWIEM - BEDZIE MOZNA SCIAGANC JEDEN IF
                            ' ROZROZNIAJACY ASN czy RECV
                            
                            
                            ' pierwszy warunek jest dla wszystkich bo w ogole sprawdzam co mam zrobic z dopasowanym pusem w nazwie
                            If Trim(CStr(.SID(Int(x)))) <> "" Then
                                If CStr(t.mName) = CStr(.SID(Int(x))) Then
                                
                                    ' sprawdzamy czy mamy jakis recving na tym pusie
                                    ' i jesli mamy to przechodzimy i patrzymy jaka jest wartosc recv
                                    If .DTYP(Int(x)) = "RECV" Then
                                    
                                        ' jesli wartosc recv jest taka sama jak pusa to znaczy ze zostal on pieknie przyjety
                                        If CLng(.qty(Int(x))) = CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                            
                                        ' druga wersja to sprawdzenie czy recv jest na zero - to trzeba jako tako zignorowac
                                        ElseIf CLng(.qty(Int(x))) = 0 Then
                                            ' NOP
                                        
                                        ' trzecia opcja to w ogole przyjecie pusa na wartosc inna niz wartosc pusa
                                        ElseIf CLng(.qty(Int(x))) <> CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                        
                                        End If
                                    ElseIf .DTYP(Int(x)) = "ASN" Then
                                        
                                        ' to jest zsynchronizowane z recv na zero
                                        ' tyle ze samo recv na zero to mozna ignorowac
                                        ' natomiast asn dr jest na potrzeby dopisania gdy transport nie przychodzi do fabyrki
                                        ' wartoscia zgodna z tym co bylo na pickup sheet
                                        If CLng(.qty(Int(x))) = CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                         ' druga wersja to sprawdzenie czy recv jest na zero - to trzeba jako tako zignorowac
                                        ElseIf CLng(.qty(Int(x))) = 0 Then
                                            ' NOP
                                        
                                        ' trzecia opcja to w ogole przyjecie pusa na wartosc inna niz wartosc pusa
                                        ElseIf CLng(.qty(Int(x))) <> CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                        
                                        End If
                                    End If
                                End If
                            End If
                        Next t
                    Else
                        Exit Sub
                    End If
                Next x
                
                If m.getCode("ms9ph100") = CStr("R6086") Then
                    Exit Do
                End If
                
                m.sendKeys "<pf8>"
            Loop While True
        End If
    End With
    
    

End Sub


Public Sub fillCollections(ByRef notYetReceivedWithPastEdaCollection As Collection, ByRef ic As Collection)


    Dim ii As IItem

    If ppum0coll.COUNT > 0 Then
        For Each t In ppum0coll
        
        
            If CDate(t.mDeliveryDate & " " & t.mDeliveryTime) < Now Then
            
                    
                notYetReceivedWithPastEdaCollection.Add t
            Else
                For Each ii In ic
                    iidate = CStr(Format(CDate(ii.getDate), "yyyy-mm-dd"))
                    tdate = CStr(Format(CDate(t.mDeliveryDate), "yyyy-mm-dd"))
                    If iidate = tdate Then
                    
                        ii.getInTransits.Add t
                        Exit For
                    End If
                Next ii
            End If
        Next t
    End If

End Sub

Private Sub Class_Initialize()
    Set HOW_MANY_DAYS_FOR_PPUS0 = ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0")
    Set ppum0coll = New Collection
End Sub

Private Sub Class_Terminate()
    Set ppum0coll = Nothing
End Sub
