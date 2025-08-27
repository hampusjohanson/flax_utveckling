Attribute VB_Name = "DataLabels1"
Option Explicit

Sub DataLabels1()
  
    Dim pptApp As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pptChart As Object
    Dim ser As Object
    Dim i As Integer
    Dim lbl As Object
    Dim labelValue As String
    Dim excelSheet As Object
    Dim chartType As XlChartType

    ' Get PowerPoint and active slide
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.slide

    ' Find chart on slide
    For Each pptShape In pptSlide.Shapes
        If pptShape.hasChart Then
            Set pptChart = pptShape.chart
            Exit For
        End If
    Next pptShape

    If pptChart Is Nothing Then
        MsgBox "No chart found on this slide.", vbExclamation
        Exit Sub
    End If

    ' Get chart type
    chartType = pptChart.chartType

    ' Get the first series in the chart
    Set ser = pptChart.SeriesCollection(1)

    ' Get the embedded Excel sheet
    Set excelSheet = pptChart.chartData.Workbook.Sheets(1)

    ' Remove existing data labels to reset settings
    ser.HasDataLabels = False

    ' Add new data labels
    ser.HasDataLabels = True

    ' Apply leader lines and ensure they appear
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Assign label text from Excel (Column A)
        labelValue = excelSheet.Cells(i + 1, 1).value
        lbl.text = labelValue

        ' **Set label positioning based on chart type**
        Select Case chartType
            Case xlXYScatter, xlXYScatterLines, xlXYScatterSmooth
                lbl.Position = xlLabelPositionAbove ' Scatter plots work better with "Above"
            Case xlColumnClustered, xlColumnStacked, xlColumnStacked100
                lbl.Position = xlLabelPositionOutsideEnd ' Outside column bars
            Case Else
                lbl.Position = xlLabelPositionBestFit ' Default if chart type is unknown
        End Select

        ' Enable leader lines
        On Error Resume Next
        lbl.ShowLeaderLines = True ' Some versions of PowerPoint may not fully support this
        On Error GoTo 0

        ' Adjust text formatting
        With lbl
            .Font.Name = "Arial"
            .Font.size = 7
            .Font.color = RGB(17, 21, 66)
            On Error Resume Next
            .HorizontalAlignment = xlHAlignLeft ' Left-align text
            On Error GoTo 0
        End With


    Next i

    ' Force a refresh of the chart (ensures leader lines are applied)
    pptChart.Refresh

End Sub




' Funktion för att beräkna medianen (exempel)
Function Median(ser As Object) As Double
    Dim values() As Double
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double

    ' Lägg alla X-värden i en array
    ReDim values(ser.Points.count - 1)
    For i = 1 To ser.Points.count
        values(i - 1) = ser.Points(i).left ' Eller använd annan data för X-värdena
    Next i

    ' Sortera värdena
    For i = 0 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) > values(j) Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i

    ' Beräkna medianen
    If UBound(values) Mod 2 = 0 Then
        ' Om jämnt antal element, ta medelvärdet av de två mittersta
        Median = (values(UBound(values) \ 2) + values(UBound(values) \ 2 + 1)) / 2
    Else
        ' Om udda antal element, ta mittvärdet
        Median = values(UBound(values) \ 2)
    End If
End Function

