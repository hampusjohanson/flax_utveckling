Attribute VB_Name = "Sub_7"
Sub Sub_7()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xMedian As Double
    Dim adjustmentFactor As Double
    Dim rowIndex As Integer

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

    ' Hämta diagrammets Excel-datablad
    On Error Resume Next
    chartShape.chart.chartData.Activate
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo 0

    If chartDataWorkbook Is Nothing Then
        MsgBox "Kunde inte öppna diagramdata. Kontrollera kompatibilitet med macOS.", vbCritical
        Exit Sub
    End If

    Set chartSheet = chartDataWorkbook.Worksheets(1)

    ' Hämta medianvärdet och justeringsfaktor
    On Error Resume Next
    xMedian = CDbl(chartSheet.Cells(3, 21).value) ' Medianvärde (kolumn U, rad 3)
    adjustmentFactor = 0.01 * (CDbl(chartSheet.Cells(3, 19).value) - CDbl(chartSheet.Cells(3, 20).value)) ' 10% av X-intervallet
    On Error GoTo 0

    ' Lägg till dataetiketter baserat på medianvärde och justeringsfaktor
    With chartShape.chart.SeriesCollection(1)
        .ApplyDataLabels
        For rowIndex = 2 To 51
            With .Points(rowIndex - 1).dataLabel
                .text = chartSheet.Cells(rowIndex, 1).value ' Använd värde från kolumn A
                .Font.Name = "Arial"
                .Font.size = 7
                .Font.color = RGB(0, 0, 0)
                If chartSheet.Cells(rowIndex, 2).value < xMedian - adjustmentFactor Then
                    .Position = xlLabelPositionLeft
                Else
                    .Position = xlLabelPositionRight
                End If
            End With
        Next rowIndex
    End With

    ' Stäng Excel-databladsboken
    chartDataWorkbook.Close


End Sub

