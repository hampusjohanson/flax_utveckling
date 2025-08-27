Attribute VB_Name = "Lines_41"
Sub Lines_41()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim chartIndex As Integer
    Dim SeriesIndex As Integer
    Dim lineWidth As Single

    ' Set the line width to 1.75
    lineWidth = 1.75

    ' === Loop through all charts on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    chartIndex = 1 ' Track chart number

    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart

            ' Loop through all series in the chart
            For SeriesIndex = 1 To chartObject.SeriesCollection.count
                Set series = chartObject.SeriesCollection(SeriesIndex)

                ' Check if the series is visible
                If series.Format.line.visible = msoTrue Then
                    On Error Resume Next
                    With series.Format.line
                        .Weight = lineWidth ' Set the line weight to 1.75
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "Error applying line weight to visible series " & SeriesIndex & " in chart " & chartIndex & ": " & Err.Description
                        Err.Clear
                    Else
                        Debug.Print "Line weight successfully applied to visible series " & SeriesIndex & " in chart " & chartIndex
                    End If
                    On Error GoTo 0
                Else
                    Debug.Print "Series " & SeriesIndex & " in chart " & chartIndex & " is not visible, skipping."
                End If
            Next SeriesIndex

            chartIndex = chartIndex + 1 ' Move to the next chart
        End If
    Next chartShape
End Sub


