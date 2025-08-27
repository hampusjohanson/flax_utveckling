Attribute VB_Name = "Mac_LD_New_4"
Sub Ld_4()
    ' Variabler
    Dim pptSlide As slide
    Dim embeddedChart As shape
    Dim chartWorkbook As Object ' Excel arbetsboken f�r diagramdata
    Dim chartSheet As Object ' Referens till Excel-arket
    Dim shapeIndex As Integer
    Dim shapeType As Integer
    Dim selectedSeries As Object ' F�r att referera till vald serie
    Dim seriesCount As Integer
    Dim chartType As Integer

    ' Hitta diagrammet p� sliden
    Set pptSlide = ActiveWindow.View.slide

    ' S�k igenom alla objekt p� sliden f�r att hitta ett diagram
    For shapeIndex = 1 To pptSlide.Shapes.count
        shapeType = pptSlide.Shapes(shapeIndex).Type
        If shapeType = msoChart Then
            Set embeddedChart = pptSlide.Shapes(shapeIndex)
            Exit For ' Hitta och avsluta loopen
        End If
    Next shapeIndex

    ' Kontrollera om vi hittade ett diagram
    If embeddedChart Is Nothing Then
        MsgBox "Inget diagram hittades p� sliden.", vbExclamation
        Exit Sub
    End If

    ' H�mta Excel arbetsboken och arket f�r diagrammet
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

    ' L�gg till en ny serie med X-v�rden fr�n B2:B51 och Y-v�rden fr�n C2:C51
    embeddedChart.chart.SeriesCollection.newSeries
    Set selectedSeries = embeddedChart.chart.SeriesCollection(1) ' V�lj den nya serien

    ' S�tt X-v�rden och Y-v�rden fr�n Excel genom att direkt referera till omr�den i Excel-arket
    selectedSeries.xValues = chartSheet.Range("B2:B51") ' X-v�rden fr�n B2:B51
    selectedSeries.values = chartSheet.Range("C2:C51") ' Y-v�rden fr�n C2:C51

    ' H�mta diagramtyp f�r att se om det �r ett scatter-diagram
    chartType = selectedSeries.chartType

    ' Kontrollera om diagrammet �r ett scatter-diagram (xlXYScatter)
    If chartType = xlXYScatter Or chartType = xlXYScatterLines Then
        ' Justera mark�rer: Runda, storlek 5, f�rg 17,21,66
        With selectedSeries.Format.Marker
            .Style = msoMarkerCircle   ' S�tt marker som cirklar
            .size = 5                 ' S�tt marker storlek till 5
            .ForeColor.RGB = RGB(17, 21, 66)  ' F�rg: RGB(17, 21, 66)
        End With
    Else
        MsgBox "Diagrammet �r inte av typen Scatter. Mark�rer kan inte justeras."
    End If
End Sub

