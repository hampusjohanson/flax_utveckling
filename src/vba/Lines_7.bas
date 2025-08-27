Attribute VB_Name = "Lines_7"
Sub Lines_7()
    'Remove markers
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim chartIndex As Integer
    Dim SeriesIndex As Integer

    ' === Loop through all charts on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    chartIndex = 1 ' Track chart number
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart

            ' Loop through the first 12 series in the chart
            For SeriesIndex = 1 To 12
                If SeriesIndex <= chartObject.SeriesCollection.count Then
                    Set series = chartObject.SeriesCollection(SeriesIndex)

                    ' Remove the markers from the series
                    On Error Resume Next
                    series.MarkerStyle = xlMarkerStyleNone
                    If Err.Number <> 0 Then
                        Debug.Print "Error removing markers for series " & SeriesIndex & " in chart " & chartIndex & ": " & Err.Description
                        Err.Clear
                    Else
                        Debug.Print "Markers removed successfully for series " & SeriesIndex & " in chart " & chartIndex
                    End If
                    On Error GoTo 0
                End If
            Next SeriesIndex
            chartIndex = chartIndex + 1
        End If
    Next chartShape
End Sub

