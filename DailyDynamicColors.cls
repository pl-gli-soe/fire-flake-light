VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyDynamicColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IDynamicColors


Private ash As Worksheet
Private flagaLiczenia As Boolean
Private zasiegDoNadaniaDynamicznychKolorow As Range
Private r As Range


' colors
Private redColor As Variant
Private pinkColor As Variant
Private blackColor As Variant
Private yellowColor As Variant
Private greyColor As Variant


Private Sub Class_Initialize()

    If CStr(ThisWorkbook.FullName) = CStr(ActiveWorkbook.FullName) Then
        Set ash = ActiveSheet
        flagaLiczenia = True
        
        redColor = ThisWorkbook.Sheets("register").Range("red").Interior.Color
        pinkColor = ThisWorkbook.Sheets("register").Range("pink").Interior.Color
        blackColor = ThisWorkbook.Sheets("register").Range("black").Interior.Color
        yellowColor = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
        greyColor = ThisWorkbook.Sheets("register").Range("grey").Interior.Color
        
        
        
    Else
        Set ash = Nothing
        flagaLiczenia = False
    End If
End Sub

Private Sub IDynamicColors_assignDynamicColorsrange()

    If flagaLiczenia Then
        Set zasiegDoNadaniaDynamicznychKolorow = ash.Range("b4")
        Set zasiegDoNadaniaDynamicznychKolorow = przeliczZasieg()
    End If
    
    ' Debug.Print zasiegDoNadaniaDynamicznychKolorow.Address
End Sub

Private Function przeliczZasieg() As Range

    
    Set przeliczZasieg = Nothing
    
    Dim poczatek As Range
    Dim ost_wiersz_pierwszej_kolumny As Range
    Dim ost_kolumna As Range
    
    
    Set ost_wiersz_pierwszej_kolumny = zasiegDoNadaniaDynamicznychKolorow.End(xlDown)
    Set ost_kolumna = zasiegDoNadaniaDynamicznychKolorow.End(xlToRight)
    
    Do
        Set zasiegDoNadaniaDynamicznychKolorow = zasiegDoNadaniaDynamicznychKolorow.Offset(0, 1)
    Loop While CStr(zasiegDoNadaniaDynamicznychKolorow) <> "Past due rqms"
    
    Set poczatek = zasiegDoNadaniaDynamicznychKolorow.Offset(1, 1)
    
    Set przeliczZasieg = Range(poczatek, Cells(ost_wiersz_pierwszej_kolumny.Row, ost_kolumna.Column))
    
End Function


Private Function IDynamicColors_getDataScope() As Range
    Set IDynamicColors_getDataScope = zasiegDoNadaniaDynamicznychKolorow
End Function

Private Sub IDynamicColors_recalcColors()


    If flagaLiczenia And (Not zasiegDoNadaniaDynamicznychKolorow Is Nothing) Then
    
        For Each r In zasiegDoNadaniaDynamicznychKolorow
            If CStr(ash.Cells(4, r.Column)) = CStr("Ending Balance") Then
                
                ' pierwszy warunek spelniony to jest kolor na ebal
                ' teraz przechodzimy i sprawdzamy jaka wartosc ma ending balance i jesli ma cos konkret
                ' to dajemy odpowiednio red or pink
                
                If CLng(r) < 0 Then
                    r.Interior.Color = CLng(redColor)
                    r.Font.Color = CLng(blackColor)
                ElseIf CLng(r.Offset(0, 1)) > CLng(r) Then
                    r.Interior.Color = CLng(pinkColor)
                    r.Font.Color = CLng(blackColor)
                Else
                    r.Interior.Color = ash.Cells(4, r.Column).Interior.Color
                    r.Font.Color = ash.Cells(4, r.Column).Font.Color
                End If
            End If
        Next r
    End If
End Sub
