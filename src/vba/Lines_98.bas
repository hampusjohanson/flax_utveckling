Attribute VB_Name = "Lines_98"
Sub Lines_98()
    Dim pptSlide As slide
    Dim chart As chart
    Dim yAxis As Axis

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the first chart on the slide
    For Each shape In pptSlide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart
            Exit For
        End If
    Next shape

    ' If no chart found, show an error and exit
    If chart Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Get the Y axis (Primary vertical axis)
    Set yAxis = chart.Axes(xlValue)

    ' Set the Y axis line color to no outline/no color
    yAxis.Format.line.visible = msoFalse

    MsgBox "Y Axis line color set to no outline/no color."
End Sub

