Attribute VB_Name = "Capital_5_Add_crossings"
Sub Mac_Cap_add_crossings()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim horizontalCrossing As Double, verticalCrossing As Double

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart on the slide
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Access the chart's data workbook
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)

    ' Calculate axis crossing points using formulas in the chart's data sheet
    With chartSheet
        .Range("I2").formula = "=MIN(B2:B13)-I9"
        .Range("I3").formula = "=MAX(B2:B13)+I9"
        .Range("I4").formula = "=MIN(C2:C13)-I9"
        .Range("I5").formula = "=MAX(C2:C13)+I9"
        .Range("I6").formula = "=MEDIAN(B2:B13)"
        .Range("I7").formula = "=MEDIAN(C2:C13)"
        .Range("I9").value = 0.02
    End With

    ' Retrieve the calculated values for axis limits and crossings
    xMin = chartSheet.Range("I2").value
    xMax = chartSheet.Range("I3").value
    yMin = chartSheet.Range("I4").value
    yMax = chartSheet.Range("I5").value
    horizontalCrossing = chartSheet.Range("I6").value
    verticalCrossing = chartSheet.Range("I7").value

    ' Adjust chart axis properties
    With chartShape.chart
        ' Set axis scales and crossing points
        .Axes(xlCategory).MinimumScale = xMin
        .Axes(xlCategory).MaximumScale = xMax
        .Axes(xlCategory).CrossesAt = horizontalCrossing
        .Axes(xlValue).MinimumScale = yMin
        .Axes(xlValue).MaximumScale = yMax
        .Axes(xlValue).CrossesAt = verticalCrossing

        ' Adjust the appearance of the axis lines
        With .Axes(xlCategory).Format.line
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66) ' Line color
            .DashStyle = msoLineLongDash ' Long dashed line
            .Weight = 0.25 ' Line thickness
        End With

        With .Axes(xlValue).Format.line
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66) ' Line color
            .DashStyle = msoLineLongDash ' Long dashed line
            .Weight = 0.25 ' Line thickness
        End With
    End With

    ' Close the chart's data workbook
    chartShape.chart.chartData.Workbook.Close
End Sub

