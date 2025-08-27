Attribute VB_Name = "AA_Series_test_8"
Sub AA_Series_8()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim originalBottom As Single
    Dim newHeight As Single

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

    ' Convert cm to points (PowerPoint uses points, 1 cm ˜ 28.35 points)
    newHeight = 11.76 * 28.35

    ' Calculate the current bottom position
    originalBottom = chartShape.Top + chartShape.height

    ' Set new height while keeping bottom position fixed
    chartShape.height = newHeight
    chartShape.Top = originalBottom - newHeight ' Adjust position upwards

   
End Sub

