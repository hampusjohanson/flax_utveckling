Attribute VB_Name = "ToppA_4f"
Sub ToppA_4f()
    ' Define variable for the table
    Dim currentSlide As slide
    Dim shapeStronger As shape

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Try to access the long_stronger table by name
    On Error Resume Next
    Set shapeStronger = currentSlide.Shapes("long_stronger")
    On Error GoTo 0 ' Reset error handling

    ' Check if the shape exists and is a table
    If Not shapeStronger Is Nothing Then
        ' Change position by setting Left (X) and Top (Y) properties
        shapeStronger.left = 8.68 * 28.35 ' X position in cm converted to points
        shapeStronger.Top = 5.81 * 28.35  ' Y position in cm converted to points

        Debug.Print "'long_stronger' table moved to new position: (" & shapeStronger.left & ", " & shapeStronger.Top & ")"
    Else
        MsgBox "'long_stronger' table not found on the current slide."
    End If
End Sub

