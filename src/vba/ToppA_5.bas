Attribute VB_Name = "ToppA_5"
' Macro: Remove all borders from long_stronger and long_weaker tables
Sub ToppA_5()
    Dim currentSlide As slide
    Dim longStrongerTable As table
    Dim longWeakerTable As table
    Dim i As Integer, j As Integer
    Dim cell As cell

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Remove borders from long_stronger table
    On Error Resume Next
    Set longStrongerTable = currentSlide.Shapes("long_stronger").table
    On Error GoTo 0

    If Not longStrongerTable Is Nothing Then
        For i = 1 To longStrongerTable.Rows.count
            For j = 1 To longStrongerTable.Columns.count
                Set cell = longStrongerTable.cell(i, j)

                ' Completely remove all borders
                With cell.Borders(ppBorderTop)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderLeft)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderRight)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderBottom)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderDiagonalDown)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderDiagonalUp)
                    .visible = msoFalse
                End With

                ' Ensure fill is transparent
                cell.shape.Fill.Transparency = 1
            Next j
        Next i
    Else
        MsgBox "Table 'long_stronger' not found.", vbExclamation
    End If

    ' Remove borders from long_weaker table
    On Error Resume Next
    Set longWeakerTable = currentSlide.Shapes("long_weaker").table
    On Error GoTo 0

    If Not longWeakerTable Is Nothing Then
        For i = 1 To longWeakerTable.Rows.count
            For j = 1 To longWeakerTable.Columns.count
                Set cell = longWeakerTable.cell(i, j)

                ' Completely remove all borders
                With cell.Borders(ppBorderTop)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderLeft)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderRight)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderBottom)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderDiagonalDown)
                    .visible = msoFalse
                End With
                With cell.Borders(ppBorderDiagonalUp)
                    .visible = msoFalse
                End With

                ' Ensure fill is transparent
                cell.shape.Fill.Transparency = 1
            Next j
        Next i
    Else
        MsgBox "Table 'long_weaker' not found.", vbExclamation
    End If

End Sub

