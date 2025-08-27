Attribute VB_Name = "ToppA_6"
' Macro: ToppA_7 - Remove all borders from long_stronger and long_weaker tables
Sub ToppA_6()
    Dim currentSlide As slide
    Dim longStrongerTable As table
    Dim longWeakerTable As table
    Dim cell As cell
    Dim i As Integer, j As Integer

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Access the long_stronger table
    On Error Resume Next
    Set longStrongerTable = currentSlide.Shapes("long_stronger").table
    On Error GoTo 0

    If Not longStrongerTable Is Nothing Then
        ' Loop through all cells in the table
        For i = 1 To longStrongerTable.Rows.count
            For j = 1 To longStrongerTable.Columns.count
                Set cell = longStrongerTable.cell(i, j)

                ' Remove all borders for the current cell
                With cell.Borders(ppBorderTop)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderLeft)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderRight)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderBottom)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderDiagonalDown)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderDiagonalUp)
                    .Transparency = 1
                End With

                ' Ensure fill is transparent
                cell.shape.Fill.Transparency = 1
            Next j
        Next i
    Else
        MsgBox "Table 'long_stronger' not found.", vbExclamation
    End If

    ' Access the long_weaker table
    On Error Resume Next
    Set longWeakerTable = currentSlide.Shapes("long_weaker").table
    On Error GoTo 0

    If Not longWeakerTable Is Nothing Then
        ' Loop through all cells in the table
        For i = 1 To longWeakerTable.Rows.count
            For j = 1 To longWeakerTable.Columns.count
                Set cell = longWeakerTable.cell(i, j)

                ' Remove all borders for the current cell
                With cell.Borders(ppBorderTop)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderLeft)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderRight)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderBottom)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderDiagonalDown)
                    .Transparency = 1
                End With
                With cell.Borders(ppBorderDiagonalUp)
                    .Transparency = 1
                End With

                ' Ensure fill is transparent
                cell.shape.Fill.Transparency = 1
            Next j
        Next i
    Else
        MsgBox "Table 'long_weaker' not found.", vbExclamation
    End If


End Sub
