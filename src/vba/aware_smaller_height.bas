Attribute VB_Name = "aware_smaller_height"
Sub aware_smaller_height()
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
    newHeight = 10.24 * 28.35

    ' Calculate the current bottom position
    originalBottom = chartShape.Top + chartShape.height

    ' Set new height while keeping bottom position fixed
    chartShape.height = newHeight
    chartShape.Top = originalBottom - newHeight ' Adjust position upwards

   
End Sub

