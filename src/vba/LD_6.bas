Attribute VB_Name = "LD_6"

Sub Mac_LD_5()
   'Rensa allt pÂ bold
   
   
    Dim pptSlide As slide
    Dim rightmostTableShape As shape
    Dim rightmostTable As table
    Dim i As Integer, j As Integer
    Dim rightMostPosition As Single

    ' === Hitta tabellen längst till höger på sliden ===
    Set pptSlide = ActiveWindow.View.slide
    rightMostPosition = -99999 ' Initialt lågt värde för att hitta högerpositionen
    Set rightmostTableShape = Nothing

    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            If shape.left > rightMostPosition Then
                rightMostPosition = shape.left
                Set rightmostTableShape = shape
            End If
        End If
    Next shape

    If rightmostTableShape Is Nothing Then
        MsgBox "Ingen tabell hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' === Hämta referens till tabellen längst till höger ===
    Set rightmostTable = rightmostTableShape.table

    ' === Iterera över alla celler och ta bort fetmarkering ===
    For i = 1 To rightmostTable.Rows.count
        For j = 1 To rightmostTable.Columns.count
            With rightmostTable.cell(i, j).shape.TextFrame.textRange.Font
                .Bold = msoFalse
            End With
        Next j
    Next i

End Sub

