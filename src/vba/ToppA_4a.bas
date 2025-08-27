Attribute VB_Name = "ToppA_4a"
Sub ToppA_4a()
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
    Set shapeStronger = currentSlide.Shapes("long_stronger")
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

