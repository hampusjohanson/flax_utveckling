Attribute VB_Name = "Module17"
Sub test_2()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObj As chart
    Dim seriesCount As Integer
    Dim numCategories As Integer
    Dim i As Integer, j As Integer
    Dim visibleSeries(1 To 2) As series
    Dim visibleSeriesCount As Integer
    Dim dataPointValue As Double

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart named "Awareness"
    Set chartShape = Nothing
    For Each chartShape In pptSlide.Shapes
        If chartShape.Name = "Awareness" Then
            Set chartObj = chartShape.chart
            Exit For
        End If
    Next chartShape

    ' Check if the chart is found
    If chartShape Is Nothing Then
        MsgBox "No chart named 'Awareness' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Get total number of series
    seriesCount = chartObj.SeriesCollection.count
    Debug.Print "Total series count: " & seriesCount

    ' Identify the first two visible series
    visibleSeriesCount = 0
    For j = 1 To seriesCount
        If chartObj.SeriesCollection(j).Format.Fill.visible Then
            visibleSeriesCount = visibleSeriesCount + 1
            Set visibleSeries(visibleSeriesCount) = chartObj.SeriesCollection(j)
            If visibleSeriesCount = 2 Then Exit For
        End If
    Next j

    ' If fewer than 2 visible series are found
    If visibleSeriesCount < 2 Then
        MsgBox "Less than two visible series found.", vbExclamation
        Exit Sub
    End If

    ' Determine the number of categories (assumes all series have the same number of points)
    numCategories = visibleSeries(1).Points.count
    Debug.Print "Categories (X-axis count): " & numCategories

    ' Ensure there's at least one category
    If numCategories = 0 Then
        MsgBox "No categories found in the chart.", vbExclamation
        Exit Sub
    End If

    ' Print out the data point values for the first and second visible series
    Debug.Print "Data Points of First Two Visible Series:"
    For i = 1 To numCategories
        Debug.Print "Category " & i & ": " & _
                    "Series 1 -> " & Format(visibleSeries(1).values(i), "0.00%") & _
                    " | Series 2 -> " & Format(visibleSeries(2).values(i), "0.00%")
    Next i

    Debug.Print "? test_2 executed correctly!"
End Sub


