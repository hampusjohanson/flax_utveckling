Attribute VB_Name = "Delete_5"
Sub Delete_5()
    Dim sectionIndex As Integer

    ' Check if there are sections in the presentation
    If ActivePresentation.SectionProperties.count = 0 Then
        MsgBox "No sections to remove in the presentation.", vbInformation, "No Sections Found"
        Exit Sub
    End If

    ' Remove sections from last to first, retaining slides
    For sectionIndex = ActivePresentation.SectionProperties.count To 1 Step -1
        ActivePresentation.SectionProperties.Delete sectionIndex, True ' Keep slides intact
    Next sectionIndex

    MsgBox "All sections have been removed successfully.", vbInformation, "Sections Removed"
End Sub

