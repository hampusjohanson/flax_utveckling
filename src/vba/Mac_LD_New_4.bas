Attribute VB_Name = "Mac_LD_New_4"
Sub Ld_4()
    ' Variabler
    Dim pptSlide As slide
    Dim embeddedChart As shape
    Dim chartWorkbook As Object ' Excel arbetsboken för diagramdata
    Dim chartSheet As Object ' Referens till Excel-arket
    Dim shapeIndex As Integer
    Dim shapeType As Integer
    Dim selectedSeries As Object ' För att referera till vald serie
    Dim seriesCount As Integer
    Dim chartType As Integer

    ' Hitta diagrammet på sliden
    Set pptSlide = ActiveWindow.View.slide

    ' Sök igenom alla objekt på sliden för att hitta ett diagram
    For shapeIndex = 1 To pptSlide.Shapes.count
        shapeType = pptSlide.Shapes(shapeIndex).Type
        If shapeType = msoChart Then
            Set embeddedChart = pptSlide.Shapes(shapeIndex)
            Exit For ' Hitta och avsluta loopen
        End If
    Next shapeIndex

    ' Kontrollera om vi hittade ett diagram
    If embeddedChart Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' Hämta Excel arbetsboken och arket för diagrammet
    Set chartWorkbook = embeddedChart.chart.chartData.Workbook
    Set chartSheet = chartWorkbook.Sheets(1)

    ' Kontrollera antalet serier innan vi raderar dem
    seriesCount = embeddedChart.chart.SeriesCollection.count

    ' Ta bort alla nuvarande serier om de finns
    If seriesCount > 0 Then
        Do While embeddedChart.chart.SeriesCollection.count > 0
            embeddedChart.chart.SeriesCollection(1).Delete
        Loop
    End If

    ' Lägg till en ny serie med X-värden från B2:B51 och Y-värden från C2:C51
    embeddedChart.chart.SeriesCollection.newSeries
    Set selectedSeries = embeddedChart.chart.SeriesCollection(1) ' Välj den nya serien

    ' Sätt X-värden och Y-värden från Excel genom att direkt referera till områden i Excel-arket
    selectedSeries.xValues = chartSheet.Range("B2:B51") ' X-värden från B2:B51
    selectedSeries.values = chartSheet.Range("C2:C51") ' Y-värden från C2:C51

    ' Hämta diagramtyp för att se om det är ett scatter-diagram
    chartType = selectedSeries.chartType

    ' Kontrollera om diagrammet är ett scatter-diagram (xlXYScatter)
    If chartType = xlXYScatter Or chartType = xlXYScatterLines Then
        ' Justera markörer: Runda, storlek 5, färg 17,21,66
        With selectedSeries.Format.Marker
            .Style = msoMarkerCircle   ' Sätt marker som cirklar
            .size = 5                 ' Sätt marker storlek till 5
            .ForeColor.RGB = RGB(17, 21, 66)  ' Färg: RGB(17, 21, 66)
        End With
    Else
        MsgBox "Diagrammet är inte av typen Scatter. Markörer kan inte justeras."
    End If
End Sub

