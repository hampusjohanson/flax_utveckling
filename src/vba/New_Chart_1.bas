Attribute VB_Name = "New_Chart_1"
Sub CreateScatterChart()
    Dim pptSlide As slide
    Dim pptShape As shape
    Dim pptChart As chart
    Dim pptDataSheet As chartData
    Dim excelWorkbook As Object
    Dim excelSheet As Object

    ' Reference the active slide
    Set pptSlide = Application.ActiveWindow.View.slide

    ' Add a new chart to the slide
    Set pptShape = pptSlide.Shapes.AddChart2(-1, xlXYScatter, 100, 100, 400, 300)
    Set pptChart = pptShape.chart

    ' Access the chart's data sheet
    Set pptDataSheet = pptChart.chartData
    pptDataSheet.Activate

    ' Get the underlying Excel workbook and sheet
    Set excelWorkbook = pptDataSheet.Workbook
    Set excelSheet = excelWorkbook.Worksheets(1)

    ' Clear existing data
    excelSheet.Cells.Clear

    ' Add sample data to the Excel sheet
    excelSheet.Cells(1, 1).value = "X Values"
    excelSheet.Cells(1, 2).value = "Y Values"
    excelSheet.Cells(2, 1).value = 1
    excelSheet.Cells(2, 2).value = 10
    excelSheet.Cells(3, 1).value = 2
    excelSheet.Cells(3, 2).value = 20
    excelSheet.Cells(4, 1).value = 3
    excelSheet.Cells(4, 2).value = 30

    ' Close the Excel workbook (but keep the data linked)
    excelWorkbook.Close

    ' Customize the chart (optional)
    pptChart.ChartTitle.text = "Sample Scatter Chart"
    pptChart.Axes(xlCategory).HasTitle = True
    pptChart.Axes(xlCategory).AxisTitle.text = "X-Axis"
    pptChart.Axes(xlValue).HasTitle = True
    pptChart.Axes(xlValue).AxisTitle.text = "Y-Axis"

    ' Inform the user
    MsgBox "Scatter chart created successfully!"

End Sub

