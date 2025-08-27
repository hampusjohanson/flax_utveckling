Attribute VB_Name = "Delete_6"
Sub Delete_6()
    Dim ppt As Presentation
    Set ppt = ActivePresentation

    ' Check if there are sections
    If ppt.SectionProperties.count > 1 Then
        Dim i As Integer
        ' Remove all sections from the last to the first
        For i = ppt.SectionProperties.count To 1 Step -1
            ppt.SectionProperties.Delete i, False ' False ensures slides are not deleted
        Next i
    End If

    MsgBox "All sections removed. The presentation now has a single section.", vbInformation
End Sub

