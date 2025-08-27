Attribute VB_Name = "AA_Series_test_9"
Sub AA_Series_9()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim newTop As Single

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart named "Awareness"
    For Each chartShape In pptSlide.Shapes
        If chartShape.Name = "Awareness" Then Exit For
    Next chartShape

    ' Check if the chart is found
    If chartShape Is Nothing Then
        MsgBox "No chart named 'Awareness' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Convert cm to points (1 cm ˜ 28.35 points)
    newTop = 5.23 * 28.35

    ' Set new vertical position
    chartShape.Top = newTop

   
End Sub

