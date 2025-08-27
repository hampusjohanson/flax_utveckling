Attribute VB_Name = "ToppA_2"
Sub ToppA_2()
    Dim currentSlide As slide
    Dim strongerTable As table
    Dim newTable As table
    Dim rowCount As Integer
    Dim firstRowHeight As Single
    Dim otherRowHeight As Single
    Dim columnWidth As Single
    Dim positionLeft As Single
    Dim positionTop As Single

    ' Dimensions and positions (converted from cm to points: 1 cm = 28.35 points)
    firstRowHeight = 0.56 * 28.35
    otherRowHeight = 0.56 * 28.35
    columnWidth = 0.9 * 28.35
    positionLeft = 7.93 * 28.35
    positionTop = 5.55 * 28.35

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Find the Stronger table
    On Error Resume Next
    Set strongerTable = currentSlide.Shapes("Stronger").table
    On Error GoTo 0

    If strongerTable Is Nothing Then
        MsgBox "The table named 'Stronger' was not found.", vbExclamation
        Exit Sub
    End If

    ' Calculate the row count for the new table (Stronger row count minus 1)
    rowCount = strongerTable.Rows.count - 1
    If rowCount < 1 Then
        MsgBox "The Stronger table does not have enough rows to create a new table.", vbExclamation
        Exit Sub
    End If

    ' Create the new table
    Dim shape As shape
    Set shape = currentSlide.Shapes.AddTable(rowCount, 1, positionLeft, positionTop, columnWidth, firstRowHeight)
    shape.Name = "border_table"
    Set newTable = shape.table

    ' Set the dimensions of the rows and columns
    Dim i As Integer, j As Integer
    Dim cell As cell
    For i = 1 To newTable.Rows.count
        ' Adjust row heights
        If i = 1 Then
            newTable.Rows(i).height = firstRowHeight
        Else
            newTable.Rows(i).height = otherRowHeight
        End If

        For j = 1 To newTable.Columns.count
            Set cell = newTable.cell(i, j)

            ' Set font size to 2 for all cells and remove fill color
            With cell.shape.TextFrame.textRange
                .Font.size = 2
            End With
            cell.shape.Fill.Transparency = 1 ' Remove fill color

            ' Remove all borders
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

            ' Set top, left, and right borders to light gray (#F2F2F2)
            With cell.Borders(ppBorderTop)
                .Weight = 0.25
                .ForeColor.RGB = RGB(242, 242, 242) ' Light gray color
                .visible = msoTrue
            End With
            With cell.Borders(ppBorderLeft)
                .Weight = 0.25
                .ForeColor.RGB = RGB(242, 242, 242) ' Light gray color
                .visible = msoTrue
            End With
            With cell.Borders(ppBorderRight)
                .Weight = 0.25
                .ForeColor.RGB = RGB(242, 242, 242) ' Light gray color
                .visible = msoTrue
            End With

            ' Add underborder to all cells
            With cell.Borders(ppBorderBottom)
                .Weight = 0.25
                .ForeColor.RGB = RGB(17, 21, 66) ' Set color to RGB(17, 21, 66)
                .visible = msoTrue
            End With
        Next j
    Next i

    ' Apply bottom borders specifically to border_table
    Dim selectedTable As table
    Set selectedTable = currentSlide.Shapes("border_table").table

    For i = 1 To selectedTable.Rows.count
        For j = 1 To selectedTable.Columns.count
            Set cell = selectedTable.cell(i, j)

            ' Add underborder to each cell
            With cell.Borders(ppBorderBottom)
                .Weight = 0.25
                .ForeColor.RGB = RGB(17, 21, 66) ' Set color to RGB(17, 21, 66)
                .visible = msoTrue
            End With
        Next j
    Next i
End Sub

