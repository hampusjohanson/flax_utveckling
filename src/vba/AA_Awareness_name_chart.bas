Attribute VB_Name = "AA_Awareness_name_chart"
Sub AA_Awareness_name_chart()
    On Error Resume Next ' Tyst felhantering

    Dim pptSlide As slide
    Dim chartShape As shape

    Set pptSlide = ActiveWindow.View.slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            chartShape.Name = "Awareness"
            Exit Sub
        End If
    Next chartShape

End Sub

