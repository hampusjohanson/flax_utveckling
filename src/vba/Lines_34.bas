Attribute VB_Name = "Lines_34"
Sub Lines_34()
'Rename Left Right
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim leftChart As shape
    Dim rightChart As shape
    Dim SlideShapes As Shapes
    Dim shapePosition As Double
    Dim minX As Double
    Dim maxX As Double

    ' Get the active slide
    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    On Error GoTo 0

    If pptSlide Is Nothing Then
        MsgBox "No active slide found. Please make sure you're in Normal View and try again.", vbExclamation
        Exit Sub
    End If

    Set SlideShapes = pptSlide.Shapes

    ' Initialize variables to identify leftmost and rightmost charts
    minX = 99999 ' Arbitrarily large value
    maxX = -99999 ' Arbitrarily small value

    ' Loop through all shapes on the slide
    For Each chartShape In SlideShapes
        If chartShape.hasChart Then ' Check if the shape contains a chart
            shapePosition = chartShape.left ' Get the left position of the shape

            ' Identify the leftmost chart
            If shapePosition < minX Then
                minX = shapePosition
                Set leftChart = chartShape
            End If

            ' Identify the rightmost chart
            If shapePosition > maxX Then
                maxX = shapePosition
                Set rightChart = chartShape
            End If
        End If
    Next chartShape

    ' Check if charts were found
    If leftChart Is Nothing Or rightChart Is Nothing Then
        MsgBox "No charts found on this slide.", vbExclamation
        Exit Sub
    End If

    ' Rename the charts
    On Error Resume Next
    leftChart.Name = "left_chart"
    rightChart.Name = "right_chart"
    On Error GoTo 0

   
End Sub


