Attribute VB_Name = "Lines_42"
Sub Lines_42()
    Dim pptSlide As slide
    Dim chartShape As shape

    ' Debug print: Start execution
    Debug.Print "--- Removing Only Marker Borders on Current Slide ---"

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Active Slide: " & pptSlide.SlideIndex

    ' Loop through all shapes on the slide
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Debug.Print "Processing chart: " & shape.Name
            
            ' Loop through all series in the chart
            With chartShape.chart
                Dim i As Integer
                For i = 1 To .SeriesCollection.count
                    With .SeriesCollection(i)
                        .MarkerStyle = xlMarkerStyleCircle
                        .MarkerSize = 6
                        .MarkerForegroundColorIndex = xlNone
                        .MarkerBackgroundColorIndex = xlNone
                    End With
                    Debug.Print "Updated markers for series: " & i
                Next i
            End With
        End If
    Next shape
    On Error GoTo 0

    ' Debug print: End execution
    Debug.Print "--- Marker Borders Removed ---"

End Sub

