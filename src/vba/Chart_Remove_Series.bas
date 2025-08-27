Attribute VB_Name = "Chart_Remove_Series"
Sub Chart_Remove_series()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim seriesCount As Integer
    Dim i As Integer

    ' Get the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart on the current slide
    On Error Resume Next
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Exit For
        End If
    Next chartShape
    On Error GoTo 0

    ' If no chart is found, exit the macro
    If chartShape Is Nothing Then
        MsgBox "No chart found on the current slide.", vbCritical
        Exit Sub
    End If

    ' Get the number of series in the chart
    seriesCount = chartShape.chart.SeriesCollection.count

    ' Loop through each series and delete it
    For i = seriesCount To 1 Step -1
        chartShape.chart.SeriesCollection(i).Delete
    Next i
End Sub

