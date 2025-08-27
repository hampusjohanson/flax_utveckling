Attribute VB_Name = "Sub_6"
Sub Sub_6()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xMax As Double, xMin As Double, yMax As Double, yMin As Double
    Dim verticalCrossing As Double, horizontalCrossing As Double

    ' Hitta aktiv slide och diagrammet
    Set pptSlide = ActiveWindow.View.slide
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape

    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' Hämta värden från diagrammets Excel-datablad
    On Error Resume Next
    chartShape.chart.chartData.Activate
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo 0

    If chartDataWorkbook Is Nothing Then
        MsgBox "Kunde inte öppna diagramdata. Kontrollera kompatibilitet med macOS.", vbCritical
        Exit Sub
    End If

    Set chartSheet = chartDataWorkbook.Worksheets(1)

    On Error Resume Next
    xMax = CDbl(chartSheet.Cells(3, 19).value) ' Max X (kolumn S, rad 3)
    xMin = CDbl(chartSheet.Cells(3, 20).value) ' Min X (kolumn T, rad 3)
    yMax = CDbl(chartSheet.Cells(4, 19).value) ' Max Y (kolumn S, rad 4)
    yMin = CDbl(chartSheet.Cells(4, 20).value) ' Min Y (kolumn T, rad 4)
    verticalCrossing = CDbl(chartSheet.Cells(6, 19).value) ' Vertikal korsning (kolumn S, rad 6)
    horizontalCrossing = CDbl(chartSheet.Cells(7, 19).value) ' Horisontell korsning (kolumn S, rad 7)
    On Error GoTo 0

    ' Uppdatera diagrammets axlar
    With chartShape.chart
        ' Uppdatera X-axel
        .Axes(xlCategory).MinimumScale = xMin
        .Axes(xlCategory).MaximumScale = xMax
        .Axes(xlCategory).CrossesAt = horizontalCrossing

        ' Uppdatera Y-axel
        .Axes(xlValue).MinimumScale = yMin
        .Axes(xlValue).MaximumScale = yMax
        .Axes(xlValue).CrossesAt = verticalCrossing
    End With

    ' Stäng Excel-datablad
    chartDataWorkbook.Close


End Sub

