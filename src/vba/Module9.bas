Attribute VB_Name = "Module9"
Sub AdjustFirstSeriesLineWeight()

    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim series As series
    Dim lineWidth As Single
    
    ' Prompt the user for the line width
    lineWidth = 7 ' Set the desired line width (can change this if needed)
    
    ' Get the active slide
    Set pptSlide = ActivePresentation.Slides(1)
    
    ' Loop through all shapes on the slide to find charts
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chart = shape.chart
            
            ' Ensure the chart has at least one series
            If chart.SeriesCollection.count > 0 Then
                ' Target the first series
                Set series = chart.SeriesCollection(1)
                
                ' Check if the chart is a line chart or scatter chart with lines
                If series.chartType = xlLine Or series.chartType = xlXYScatterLines Then
                    ' Ensure the line is visible and set the line weight
                    series.Format.line.visible = msoTrue ' Ensure line is visible
                    series.Format.line.ForeColor.RGB = RGB(0, 0, 255) ' Set line color to blue
                    series.Format.line.Weight = lineWidth ' Set the line weight
                    
                    MsgBox "Line weight set to " & lineWidth & " for the first series."
                Else
                    MsgBox "The first series is not a line chart or scatter line chart."
                End If
            Else
                MsgBox "No series found in the chart."
            End If
        End If
    Next shape
    
    MsgBox "Line weight adjustment complete."
End Sub

