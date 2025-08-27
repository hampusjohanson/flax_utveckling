Attribute VB_Name = "Topp_2"
Sub Topp_2()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim borderColor As Long
    Dim widestTableShape As shape
    Dim maxWidth As Single
    Dim shape As shape
    Dim columnPairs As Variant
    Dim pair As Variant

    ' Hitta den aktuella sliden
    Set pptSlide = ActiveWindow.View.slide


    ' Hitta tabellen "TARGET"
    On Error Resume Next
    Set targetShape = pptSlide.Shapes("TARGET")
    On Error GoTo 0

    If targetShape Is Nothing Or Not targetShape.HasTable Then
        MsgBox "Kunde inte hitta tabellen 'TARGET'. Kontrollera att tabellen finns på sliden.", vbCritical
        Exit Sub
    End If

    ' Hämta tabellen
    Set targetTable = targetShape.table

    ' Ta bort alla rader från rad 12 och nedåt
    While targetTable.Rows.count > 11
        targetTable.Rows(targetTable.Rows.count).Delete
    Wend

    ' Lägg till borders för första raden (kolumnpar 1+2, 4+5, 7+8, 10+11)
    borderColor = RGB(17, 21, 66) ' Färgkod för kantlinjen
    columnPairs = Array(Array(1, 2), Array(4, 5), Array(7, 8), Array(10, 11))

    For Each pair In columnPairs
        For colIndex = pair(0) To pair(1)
            With targetTable.cell(1, colIndex).Borders
                .Item(ppBorderLeft).visible = msoTrue
                .Item(ppBorderLeft).ForeColor.RGB = borderColor
                .Item(ppBorderLeft).Weight = 0.25

                .Item(ppBorderRight).visible = msoTrue
                .Item(ppBorderRight).ForeColor.RGB = borderColor
                .Item(ppBorderRight).Weight = 0.25

                .Item(ppBorderTop).visible = msoTrue
                .Item(ppBorderTop).ForeColor.RGB = borderColor
                .Item(ppBorderTop).Weight = 0.25

                .Item(ppBorderBottom).visible = msoTrue
                .Item(ppBorderBottom).ForeColor.RGB = borderColor
                .Item(ppBorderBottom).Weight = 0.25
            End With
        Next colIndex
    Next pair

    ' Lägg till borders och numrering i kolumnerna för rad 2 till 11
    For rowIndex = 2 To 11
        ' Lägg till numrering i specifika kolumner
        targetTable.cell(rowIndex, 1).shape.TextFrame.textRange.text = (rowIndex - 1) & "."
        targetTable.cell(rowIndex, 4).shape.TextFrame.textRange.text = (rowIndex + 9) & "."
        targetTable.cell(rowIndex, 7).shape.TextFrame.textRange.text = (rowIndex + 19) & "."
        targetTable.cell(rowIndex, 10).shape.TextFrame.textRange.text = (rowIndex + 29) & "."

        ' Borders för kolumn 1
        With targetTable.cell(rowIndex, 1).Borders
            .Item(ppBorderLeft).visible = msoTrue
            .Item(ppBorderLeft).ForeColor.RGB = borderColor
            .Item(ppBorderLeft).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 2
        With targetTable.cell(rowIndex, 2).Borders
            .Item(ppBorderRight).visible = msoTrue
            .Item(ppBorderRight).ForeColor.RGB = borderColor
            .Item(ppBorderRight).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 4
        With targetTable.cell(rowIndex, 4).Borders
            .Item(ppBorderLeft).visible = msoTrue
            .Item(ppBorderLeft).ForeColor.RGB = borderColor
            .Item(ppBorderLeft).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 5
        With targetTable.cell(rowIndex, 5).Borders
            .Item(ppBorderRight).visible = msoTrue
            .Item(ppBorderRight).ForeColor.RGB = borderColor
            .Item(ppBorderRight).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 7
        With targetTable.cell(rowIndex, 7).Borders
            .Item(ppBorderLeft).visible = msoTrue
            .Item(ppBorderLeft).ForeColor.RGB = borderColor
            .Item(ppBorderLeft).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 8
        With targetTable.cell(rowIndex, 8).Borders
            .Item(ppBorderRight).visible = msoTrue
            .Item(ppBorderRight).ForeColor.RGB = borderColor
            .Item(ppBorderRight).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 10
        With targetTable.cell(rowIndex, 10).Borders
            .Item(ppBorderLeft).visible = msoTrue
            .Item(ppBorderLeft).ForeColor.RGB = borderColor
            .Item(ppBorderLeft).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With

        ' Borders för kolumn 11
        With targetTable.cell(rowIndex, 11).Borders
            .Item(ppBorderRight).visible = msoTrue
            .Item(ppBorderRight).ForeColor.RGB = borderColor
            .Item(ppBorderRight).Weight = 0.25
            .Item(ppBorderBottom).visible = msoTrue
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With
    Next rowIndex

    ' Rensa text manuellt i specificerade kolumner
    For rowIndex = 2 To 11
        targetTable.cell(rowIndex, 2).shape.TextFrame.textRange.text = ""
        targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.text = ""
        targetTable.cell(rowIndex, 8).shape.TextFrame.textRange.text = ""
        targetTable.cell(rowIndex, 11).shape.TextFrame.textRange.text = ""
    Next rowIndex

End Sub


