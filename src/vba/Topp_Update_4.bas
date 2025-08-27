Attribute VB_Name = "Topp_Update_4"
Sub Topp_Update_4()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim i As Integer
    Dim borderColor As Long

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Find the TARGET table ===
    On Error Resume Next
    Set targetShape = pptSlide.Shapes("TARGET")
    If targetShape Is Nothing Or Not targetShape.HasTable Then
        MsgBox "No table found named TARGET.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set targetTable = targetShape.table

    ' Define border color (RGB 17, 21, 66)
    borderColor = RGB(17, 21, 66)

    ' === Set borders for column 10 ===
    For i = 1 To targetTable.Rows.count
        With targetTable.cell(i, 10).Borders
            .Item(ppBorderLeft).ForeColor.RGB = borderColor
            .Item(ppBorderLeft).Weight = 0.25
            .Item(ppBorderTop).ForeColor.RGB = borderColor
            .Item(ppBorderTop).Weight = 0.25
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With
    Next i

    ' === Set borders for column 11 ===
    For i = 1 To targetTable.Rows.count
        With targetTable.cell(i, 11).Borders
            .Item(ppBorderRight).ForeColor.RGB = borderColor
            .Item(ppBorderRight).Weight = 0.25
            .Item(ppBorderTop).ForeColor.RGB = borderColor
            .Item(ppBorderTop).Weight = 0.25
            .Item(ppBorderBottom).ForeColor.RGB = borderColor
            .Item(ppBorderBottom).Weight = 0.25
        End With
    Next i

End Sub
