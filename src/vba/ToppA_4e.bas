Attribute VB_Name = "ToppA_4e"
Sub ToppA_4e()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeStronger As shape
    Dim cell As Object ' Outer loop cell
    Dim tableCell As Object ' Inner loop cell

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_stronger table by name
    On Error Resume Next
    Set shapeStronger = currentSlide.Shapes("long_stronger")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeStronger Is Nothing Then
        ' Loop through each row in the table
        For Each cell In shapeStronger.table.Rows
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
        Debug.Print "Text formatting applied to all cells in 'long_stronger' table."
    Else
        MsgBox "'long_stronger' table not found on the current slide."
    End If
End Sub

