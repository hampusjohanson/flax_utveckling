Attribute VB_Name = "AA_Names"
Sub AA_Names()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim categoryLabels As Collection
    Dim rowIndex As Integer
    Dim currentCellValue As String
    Dim selectedLabelsCount As Integer
    Dim selectedLabels() As Variant
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

    ' Initialize collection to store selected category labels
    Set categoryLabels = New Collection

    ' Loop through each row in the table (Now checking COLUMN 2 for "Yes")
    For rowIndex = 1 To tableShape.table.Rows.count
        ' Read cell value from column 2
        currentCellValue = LCase(Trim(tableShape.table.cell(rowIndex, 2).shape.TextFrame.textRange.text))
        
        ' If the value is "yes", include corresponding A-column value from the chart
        If currentCellValue = "yes" Then
            categoryLabels.Add "A" & (rowIndex + 1) ' Store the cell reference for column A
        End If
    Next rowIndex

    ' If no rows were selected, show message and exit
    selectedLabelsCount = categoryLabels.count
    If selectedLabelsCount = 0 Then
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

    ' Convert selected category labels into an array
    ReDim selectedLabels(1 To selectedLabelsCount)
    For i = 1 To selectedLabelsCount
        selectedLabels(i) = chartSheet.Range(categoryLabels.Item(i)).value
    Next i

    ' Set the horizontal axis labels (category labels)
    With chartShape.chart.Axes(xlCategory)
        .CategoryNames = selectedLabels
    End With

    
End Sub

