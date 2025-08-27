Attribute VB_Name = "LD_6"

Sub Mac_LD_5()
   'Rensa allt p� bold
   
   
    Dim pptSlide As slide
    Dim rightmostTableShape As shape
    Dim rightmostTable As table
    Dim i As Integer, j As Integer
    Dim rightMostPosition As Single

    ' === Hitta tabellen l�ngst till h�ger p� sliden ===
    Set pptSlide = ActiveWindow.View.slide
    rightMostPosition = -99999 ' Initialt l�gt v�rde f�r att hitta h�gerpositionen
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
        MsgBox "Ingen tabell hittades p� sliden.", vbExclamation
        Exit Sub
    End If

    ' === H�mta referens till tabellen l�ngst till h�ger ===
    Set rightmostTable = rightmostTableShape.table

    ' === Iterera �ver alla celler och ta bort fetmarkering ===
    For i = 1 To rightmostTable.Rows.count
        For j = 1 To rightmostTable.Columns.count
            With rightmostTable.cell(i, j).shape.TextFrame.textRange.Font
                .Bold = msoFalse
            End With
        Next j
    Next i

End Sub

