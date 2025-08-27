Attribute VB_Name = "AA_Series_3"
Sub AA_Series_3()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim yValues As Collection
    Dim newSeries As Object
    Dim rowIndex As Integer
    Dim currentCellValue As String
    Dim selectedRowsCount As Integer
    Dim selectedValues() As Variant
    Dim i As Integer

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the table named "Brands"
    For Each tableShape In pptSlide.Shapes
        If tableShape.Name = "Brands" Then Exit For
    Next tableShape

    ' Check if table was found
    If tableShape Is Nothing Then
        MsgBox "No table named 'Brands' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Initialize collection to store selected Y-values
    Set yValues = New Collection

    ' Loop through each row in the table (Now checking COLUMN 2 for "Yes")
    For rowIndex = 1 To tableShape.table.Rows.count
        ' Read cell value from column 2
        currentCellValue = LCase(Trim(tableShape.table.cell(rowIndex, 2).shape.TextFrame.textRange.text))
        
        ' If the value is "yes", include corresponding D-column value from the chart
        If currentCellValue = "yes" Then
            yValues.Add "D" & (rowIndex + 1) ' Store row index for later retrieval
        End If
    Next rowIndex

    ' If no rows were selected, show message and exit
    selectedRowsCount = yValues.count
    If selectedRowsCount = 0 Then
        MsgBox "No rows were marked as 'Yes' in Brands.", vbExclamation
        Exit Sub
    End If

    ' Find the chart named "Awareness"
    For Each chartShape In pptSlide.Shapes
        If chartShape.Name = "Awareness" Then
            ' Access the chart's embedded Excel workbook
            Set chartDataWorkbook = chartShape.chart.chartData.Workbook
            Set chartSheet = chartDataWorkbook.Sheets(1)
            Exit For
        End If
    Next chartShape

    ' Check if the chart is found
    If chartDataWorkbook Is Nothing Then
        MsgBox "No chart named 'Awareness' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Convert selected row values into an array
    ReDim selectedValues(1 To selectedRowsCount)
    For i = 1 To selectedRowsCount
        selectedValues(i) = chartSheet.Range(yValues.Item(i)).value
    Next i

    ' Add a new series without deleting existing ones
    Set newSeries = chartShape.chart.SeriesCollection.newSeries
    newSeries.values = selectedValues

    ' Optionally, set the series name
    newSeries.Name = "Serie3" & chartShape.chart.SeriesCollection.count

    ' Apply percentage formatting to data labels
    newSeries.ApplyDataLabels
    newSeries.DataLabels.NumberFormat = "0%" ' Format as percentage

    ' Set the data labels to white
    With newSeries
        .ApplyDataLabels
        .DataLabels.Font.color = RGB(255, 255, 255)
    End With

    ' Set the fill color of the series
    With newSeries.Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(158, 159, 177) ' Hex #9E9FB1
        .Solid
    End With

    
End Sub

