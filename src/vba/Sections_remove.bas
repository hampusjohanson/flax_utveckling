Attribute VB_Name = "Sections_remove"
Sub RemoveSectionsAndPrepareForPDF()
    Dim slide As slide
    
    ' Loop through slides and ensure no sections are present
    On Error Resume Next
    For Each slide In ActivePresentation.Slides
        ' Clear any section titles if possible (sections may not exist on Mac)
        slide.HeadersFooters.Clear
    Next slide
    On Error GoTo 0

    ' Optionally, you can adjust slide properties here if needed for better export (e.g., orientation, transitions)
    ' Example: Set slide show properties to fit for PDF export
    With ActivePresentation.PageSetup
        .slideWidth = 1024 ' Adjust as needed for your preferred size
        .slideHeight = 768 ' Adjust as needed for your preferred size
    End With

    ' The presentation is now ready for PDF export
    MsgBox "Presentation is now cleaned up and ready for PDF export.", vbInformation
End Sub

