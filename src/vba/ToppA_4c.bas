Attribute VB_Name = "ToppA_4c"
Sub ToppA_4c()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeStronger As shape

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_stronger table by its name
    On Error Resume Next
    Set shapeStronger = currentSlide.Shapes("long_stronger")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeStronger Is Nothing Then
        ' Move the table by setting its Left and Top properties
        ' These positions are in points (1 point = 1/72 of an inch)
        shapeStronger.left = 10 * 28.35  ' Change the X position (horizontal) - adjust as needed
        shapeStronger.Top = 5 * 28.35    ' Change the Y position (vertical) - adjust as needed

        Debug.Print "Table moved to new position."
    Else
        MsgBox "'long_stronger' table not found on the current slide."
    End If
End Sub

