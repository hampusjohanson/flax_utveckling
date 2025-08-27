Attribute VB_Name = "Background_1_1"
Sub Background_1_1()
    Dim slide As slide
    Dim shape As shape
    Dim table1 As shape
    Dim table2 As shape

    ' Ensure there is an active presentation
    If ActivePresentation Is Nothing Then
        MsgBox "No active presentation found.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Ensure a slide is selected
    If ActiveWindow.Selection.SlideRange.count = 0 Then
        MsgBox "Please select a slide.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Get the active slide
    Set slide = ActiveWindow.View.slide

    ' Loop through shapes to find tables
    For Each shape In slide.Shapes
        If shape.HasTable Then
            If table1 Is Nothing Then
                Set table1 = shape
            ElseIf table2 Is Nothing Then
                Set table2 = shape
                Exit For ' Stop after finding two tables
            End If
        End If
    Next shape

    ' Check if two tables were found
    If table1 Is Nothing Or table2 Is Nothing Then
        MsgBox "Could not find two tables on the slide.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Assign names based on position
    If table1.left < table2.left Then
        table1.Name = "LEFTIE"
        table2.Name = "RIGHTIE"
    Else
        table1.Name = "RIGHTIE"
        table2.Name = "LEFTIE"
    End If

End Sub

