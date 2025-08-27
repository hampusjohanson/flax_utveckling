Attribute VB_Name = "Module12"
Sub DataLabelsWithLeaderLines()
    Dim pptApp As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pptChart As Object
    Dim ser As Object
    Dim i As Integer
    Dim lbl As Object
    Dim labelValue As String
    Dim excelSheet As Object

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

    ' Get the first series in the chart
    Set ser = pptChart.SeriesCollection(1)

    ' Get the embedded Excel sheet
    Set excelSheet = pptChart.chartData.Workbook.Sheets(1)

    ' Remove existing labels, then re-add
    ser.HasDataLabels = False
    DoEvents
    ser.HasDataLabels = True
    DoEvents

    ' Apply labels and force leader lines
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Assign label text from Excel (Column A)
        labelValue = excelSheet.Cells(i + 1, 1).value
        lbl.text = labelValue

        ' **Move label away to force leader lines**
        lbl.Position = xlLabelPositionAbove

        ' **Enable leader lines**
        On Error Resume Next
        lbl.ShowLeaderLines = True
        On Error GoTo 0

        ' Adjust text formatting
        With lbl
            .Font.Name = "Arial"
            .Font.size = 7
            .Font.color = RGB(17, 21, 66)
            On Error Resume Next
            .HorizontalAlignment = xlHAlignLeft
            On Error GoTo 0
        End With

         Next i

    ' **Final refresh to enforce leader lines**
    pptChart.Refresh
End Sub

