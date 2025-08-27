Attribute VB_Name = "Text_Input_2"
Sub Text_Input_2()
    Dim pptSlide As slide
    Dim s As shape
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Iterate through all shapes on the slide
    For Each s In pptSlide.Shapes
        ' Check if the shape is named "Text_Bold" and delete it
        If s.Name = "Text_Bold" Then
            s.Delete
            Debug.Print "Text box 'Text_Bold' deleted."
            Exit Sub ' Exit after deleting, assuming only one instance
        End If
    Next s

    ' If no text box was found, print a message
    Debug.Print "No text box named 'Text_Bold' found on the slide."
End Sub

