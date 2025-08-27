Attribute VB_Name = "Chart_Add_Series"
Sub UpdateChartSeriesWithFalse()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xRange As Object, yRange As Object
    Dim xValuesArray() As Double
    Dim yValuesArray() As Double
    Dim cell As Object
    Dim i As Integer

    ' Get the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart on the current slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Exit For
        End If
    Next chartShape

    ' If no chart is found, exit the macro
    If chartShape Is Nothing Then
        MsgBox "No chart found on the current slide.", vbCritical
        Exit Sub
    End If

    ' Open the chart's Excel workbook
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)

    ' Define the source ranges
    Set xRange = chartSheet.Range("B2:B51")
    Set yRange = chartSheet.Range("C2:C51")

    ' Initialize arrays to store filtered values
    ReDim xValuesArray(1 To xRange.Rows.count)
    ReDim yValuesArray(1 To yRange.Rows.count)
    i = 1

    ' Filter valid data
    For Each cell In xRange
        ' Check for "FALSE" or invalid data
        If cell.value <> "FALSE" And IsNumeric(cell.value) Then
            xValuesArray(i) = cell.value
            yValuesArray(i) = yRange.Cells(cell.row - xRange.row + 1, 1).value ' Match Y value
            i = i + 1
        End If
    Next cell

    ' Resize arrays to fit valid data
    ReDim Preserve xValuesArray(1 To i - 1)
    ReDim Preserve yValuesArray(1 To i - 1)

    ' Remove existing series and add a new one
    With chartShape.chart
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        .SeriesCollection.newSeries
        With .SeriesCollection(1)
            .xValues = xValuesArray ' Assign filtered X values
            .values = yValuesArray  ' Assign filtered Y values
            ' Customize series appearance
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 14
            .Format.line.visible = msoFalse ' Remove line between points
        End With
    End With

    ' Close the chart's Excel workbook
    chartShape.chart.chartData.Workbook.Close

End Sub

