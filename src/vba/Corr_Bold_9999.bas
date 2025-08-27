Attribute VB_Name = "Corr_Bold_9999"
Sub Corr_Bold_9999()
    Dim pptSlide As slide
    Dim existingBox As shape
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Try to find and delete "Bold_Text" textbox
    On Error Resume Next
    Set existingBox = pptSlide.Shapes("Bold_Text")
    If Not existingBox Is Nothing Then
        existingBox.Delete
        Debug.Print "Text box 'Bold_Text' removed."
    Else
        Debug.Print "No 'Bold_Text' text box found."
    End If
    On Error GoTo 0
End Sub

