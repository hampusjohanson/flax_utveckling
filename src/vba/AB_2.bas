Attribute VB_Name = "AB_2"
Sub AB_2()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartData As Object ' To hold the ChartData Workbook
    Dim targetTable As table
    Dim cellValue As String
    Dim shape As shape

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Locate the first chart on the slide ===
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape ' Use the first chart found
            Exit For
        End If
    Next shape

    ' Check if the chart was found
    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Access the Excel data of the chart ===
    Set chartData = chartShape.chart.chartData.Workbook

    ' === Locate the "TARGET" table ===
    Set targetTable = Nothing
    For Each shape In pptSlide.Shapes
        If shape.HasTable And shape.Name = "TARGET" Then
            Set targetTable = shape.table
            Exit For
        End If
    Next shape

    ' Check if the "TARGET" table was found
    If targetTable Is Nothing Then
        MsgBox "Table 'TARGET' not found on the slide.", vbExclamation
        ' Close the chart's Excel data
        chartData.Application.Quit
        Exit Sub
    End If

    ' === Paste values into the TARGET table ===
    Dim rowIndex As Integer
    Dim colIndex As Integer

    ' Ensure the table has enough rows
    If targetTable.Rows.count < 5 Then
        MsgBox "The 'TARGET' table does not have enough rows. Add more rows to accommodate the data.", vbExclamation
        chartData.Application.Quit
        Exit Sub
    End If

    ' Row 2: B1
    rowIndex = 2
    colIndex = 2
    targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = chartData.Sheets(1).Range("B1").value

    ' Row 3: C1
    rowIndex = 3
    targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = chartData.Sheets(1).Range("C1").value

    ' Row 4: D1
    rowIndex = 4
    targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = chartData.Sheets(1).Range("D1").value

    ' Row 5: E1
    rowIndex = 5
    targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = chartData.Sheets(1).Range("E1").value

    ' Close the chart's Excel data
    chartData.Application.Quit

   
End Sub

