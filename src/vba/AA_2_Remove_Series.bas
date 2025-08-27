Attribute VB_Name = "AA_2_Remove_Series"
Sub AW_2_Remove_Series()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim seriesCount As Integer
    Dim i As Integer
    
    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide
    
    ' Loop through each shape on the slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            ' Get the number of series in the chart
            seriesCount = chartShape.chart.SeriesCollection.count
            
            ' Loop through all series and delete them
            For i = seriesCount To 1 Step -1
                chartShape.chart.SeriesCollection(i).Delete
            Next i
            
          
            Exit Sub
        End If
    Next chartShape

    ' If no chart is found on the slide
End Sub

