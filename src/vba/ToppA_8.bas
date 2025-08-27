Attribute VB_Name = "ToppA_8"
' Macro: Set the left border of all cells in 'border_table' and 'border_table_weaker' to no outline using transparency
Sub ToppA_8()
    Dim currentSlide As slide
    Dim borderTable As table
    Dim borderTableWeaker As table
    Dim cell As cell
    Dim i As Integer, j As Integer

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Access the 'border_table'
    On Error Resume Next
    Set borderTable = currentSlide.Shapes("border_table").table
    On Error GoTo 0

    If Not borderTable Is Nothing Then
        ' Loop through all cells in the table
        For i = 1 To borderTable.Rows.count
            For j = 1 To borderTable.Columns.count
                Set cell = borderTable.cell(i, j)

                ' Set the left border to transparent (no outline)
                With cell.Borders(ppBorderLeft)
                    .Transparency = 1 ' Fully transparent
                    .visible = msoTrue ' Ensure it's handled explicitly
                End With
            Next j
        Next i
    Else
        MsgBox "The table named 'border_table' was not found.", vbExclamation
    End If

    ' Access the 'border_table_weaker'
    On Error Resume Next
    Set borderTableWeaker = currentSlide.Shapes("border_table_weaker").table
    On Error GoTo 0

    If Not borderTableWeaker Is Nothing Then
        ' Loop through all cells in the table
        For i = 1 To borderTableWeaker.Rows.count
            For j = 1 To borderTableWeaker.Columns.count
                Set cell = borderTableWeaker.cell(i, j)

                ' Set the left border to transparent (no outline)
                With cell.Borders(ppBorderLeft)
                    .Transparency = 1 ' Fully transparent
                    .visible = msoTrue ' Ensure it's handled explicitly
                End With
            Next j
        Next i
    Else
        MsgBox "The table named 'border_table_weaker' was not found.", vbExclamation
    End If
End Sub

