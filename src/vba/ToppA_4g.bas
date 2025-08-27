Attribute VB_Name = "ToppA_4g"
Sub ToppA_4g()
    ' Define variable for the table and cell
    Dim currentSlide As slide
    Dim longWeakerTable As table
    Dim cell As cell
    Dim shapeWeaker As shape
    Dim rowIndex As Long

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_weaker table by its name
    On Error Resume Next
    Set shapeWeaker = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeWeaker Is Nothing Then
        If shapeWeaker.HasTable Then
            Set longWeakerTable = shapeWeaker.table
            Debug.Print "Found long_weaker table."

            ' Loop through each row and cell in the table
            For rowIndex = 1 To longWeakerTable.Rows.count
                For Each cell In longWeakerTable.Rows(rowIndex).Cells
                    If Not cell.shape.HasTextFrame Then
                        Debug.Print "Cell has no TextFrame."
                    Else
                        ' Remove fill color
                        cell.shape.Fill.BackColor.RGB = RGB(255, 255, 255) ' No fill (white)
                        cell.shape.Fill.Transparency = 1 ' Fully transparent
                        Debug.Print "Removed fill color for cell with text: " & cell.shape.TextFrame.textRange.text
                    End If
                Next cell
            Next rowIndex

            Debug.Print "Fill color removed successfully for all cells."
        Else
            MsgBox "Shape 'long_weaker' is not a table."
        End If
    Else
        MsgBox "'long_weaker' table not found on the current slide."
    End If
End Sub

Sub ToppA_4h()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeWeaker As shape

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_weaker table by its name
    On Error Resume Next
    Set shapeWeaker = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeWeaker Is Nothing Then
        ' Move the table by setting its Left and Top properties
        shapeWeaker.left = 10 * 28.35  ' Change the X position (horizontal) - adjust as needed
        shapeWeaker.Top = 5 * 28.35    ' Change the Y position (vertical) - adjust as needed

        Debug.Print "Table moved to new position."
    Else
        MsgBox "'long_weaker' table not found on the current slide."
    End If
End Sub

Sub ToppA_4i()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeWeaker As shape
    Dim cell As Object ' Outer loop cell
    Dim tableCell As Object ' Inner loop cell

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_weaker table by name
    On Error Resume Next
    Set shapeWeaker = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeWeaker Is Nothing Then
        ' Loop through each row in the table
        For Each cell In shapeWeaker.table.Rows
            ' Loop through each cell in the row
            For Each tableCell In cell.Cells
                With tableCell.shape.TextFrame.textRange
                    .Font.Name = "Arial" ' Set font to Arial
                    .Font.color.RGB = RGB(17, 21, 66) ' Set font color to RGB(17, 21, 66)
                End With
                Debug.Print "Formatted cell with text: " & tableCell.shape.TextFrame.textRange.text
            Next tableCell
        Next cell
        Debug.Print "Text formatting applied to all cells in 'long_weaker' table."
    Else
        MsgBox "'long_weaker' table not found on the current slide."
    End If
End Sub

Sub ToppA_4j()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeWeaker As shape
    Dim cell As Object ' Outer loop cell
    Dim tableCell As Object ' Inner loop cell

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_weaker table by name
    On Error Resume Next
    Set shapeWeaker = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeWeaker Is Nothing Then
        ' Loop through each row in the table
        For Each cell In shapeWeaker.table.Rows
            ' Loop through each cell in the row
            For Each tableCell In cell.Cells
                With tableCell.shape.TextFrame.textRange
                    .Font.Name = "Arial" ' Set font to Arial
                    .Font.color.RGB = RGB(17, 21, 66) ' Set font color to RGB(17, 21, 66)
                    .Font.Bold = False ' Make the text not bold
                    .Font.Italic = True ' Make the text italic
                End With
                Debug.Print "Formatted cell with text: " & tableCell.shape.TextFrame.textRange.text
            Next tableCell
        Next cell
        Debug.Print "Text formatting applied to all cells in 'long_weaker' table."
    Else
        MsgBox "'long_weaker' table not found on the current slide."
    End If
End Sub

Sub ToppA_4k()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeWeaker As shape

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_weaker table by its name
    On Error Resume Next
    Set shapeWeaker = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeWeaker Is Nothing Then
        ' Change position by setting Left (X) and Top (Y) properties
        shapeWeaker.left = 24.46 * 28.35 ' X position in cm converted to points
        shapeWeaker.Top = 5.81 * 28.35  ' Y position in cm converted to points

        Debug.Print "'long_weaker' table moved to new position: (" & shapeWeaker.left & ", " & shapeWeaker.Top & ")"
    Else
        MsgBox "'long_weaker' table not found on the current slide."
    End If
End Sub

Sub ToppA_4a_weak()
    ' Define variable for the table and cell
    Dim currentSlide As slide
    Dim longStrongerTable As table
    Dim cell As cell
    Dim shapeStronger As shape
    Dim rowIndex As Long

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_stronger table by its name
    On Error Resume Next
    Set shapeStronger = currentSlide.Shapes("long_weaker")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeStronger Is Nothing Then
        If shapeStronger.HasTable Then
            Set longStrongerTable = shapeStronger.table
            Debug.Print "Found long_stronger table."

            ' Loop through each row and cell in the table
            For rowIndex = 1 To longStrongerTable.Rows.count
                For Each cell In longStrongerTable.Rows(rowIndex).Cells
                    If Not cell.shape.HasTextFrame Then
                        Debug.Print "Cell has no TextFrame."
                    Else
                        ' Check if TextFrame exists and then set font size
                        With cell.shape.TextFrame.textRange.Font
                            .size = 7
                        End With
                        Debug.Print "Font size set to x for cell with text: " & cell.shape.TextFrame.textRange.text
                    End If
                Next cell
            Next rowIndex

            Debug.Print "Font size updated successfully for all rows."
        Else
            MsgBox "Shape 'long_stronger' is not a table."
        End If
    Else
        MsgBox "'long_stronger' table not found on the current slide."
    End If
End Sub


