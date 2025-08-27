Attribute VB_Name = "SS_3"
Sub SS_3()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xRange As Object, yRange As Object
    Dim xValuesArray() As Double
    Dim yValuesArray() As Double
    Dim cell As Object
    Dim i As Integer
    Dim relativeRow As Integer
    Dim filteredCount As Integer

    ' Get the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart on the current slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then Exit For
    Next chartShape

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

    ReDim xValuesArray(1 To xRange.Rows.count)
    ReDim yValuesArray(1 To yRange.Rows.count)
    i = 1

    ' Filter valid data
    For Each cell In xRange
        relativeRow = cell.row - xRange.Cells(1, 1).row + 1 ' 1-based offset
        If LCase(Trim(cell.value)) <> "false" And LCase(Trim(cell.value)) <> "falskt" And Trim(cell.value) <> "" Then
            If IsNumeric(cell.value) And IsNumeric(yRange.Cells(relativeRow, 1).value) Then
                xValuesArray(i) = cell.value
                yValuesArray(i) = yRange.Cells(relativeRow, 1).value
                i = i + 1
            End If
        End If
    Next cell

    filteredCount = i - 1
    If filteredCount = 0 Then
        MsgBox "No valid data found.", vbExclamation
        chartDataWorkbook.Close
        Exit Sub
    End If

    ReDim Preserve xValuesArray(1 To filteredCount)
    ReDim Preserve yValuesArray(1 To filteredCount)

    With chartShape.chart
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        .SeriesCollection.newSeries
        With .SeriesCollection(1)
            .xValues = xValuesArray
            .values = yValuesArray
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 7
            .Format.line.visible = msoFalse
        End With
    End With

    chartDataWorkbook.Close
    Debug.Print "SS_3 färdig. Antal punkter: " & filteredCount
End Sub

