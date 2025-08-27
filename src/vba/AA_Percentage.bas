Attribute VB_Name = "AA_Percentage"
Sub ChangeDataLabelsToPercentage()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartSeries As series
    Dim chartDataLabels As DataLabels

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loop through all shapes on the slide to find charts
    For Each chartShape In pptSlide.Shapes
        ' Check if the shape has a chart
        If chartShape.hasChart Then
            ' Loop through all series in the chart
            For Each chartSeries In chartShape.chart.SeriesCollection
                ' Enable data labels if not already enabled
                chartSeries.ApplyDataLabels

                ' Access the data labels of the series
                Set chartDataLabels = chartSeries.DataLabels

                ' Change the number format to percentage
                chartDataLabels.NumberFormat = "0%"
            Next chartSeries

            MsgBox "All data labels in the chart '" & chartShape.Name & "' have been changed to percentages.", vbInformation
            Exit Sub
        End If
    Next chartShape

    ' If no chart was found on the slide
    MsgBox "No chart found on the slide.", vbExclamation
End Sub

