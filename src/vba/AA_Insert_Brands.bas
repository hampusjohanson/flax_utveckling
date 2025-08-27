Attribute VB_Name = "AA_Insert_Brands"
Sub Insert_Brand_Table()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim dataRange As Object ' Use Object instead of Range for Excel data in chart
    Dim rowIndex As Integer, validRowIndex As Integer
    Dim cellValue As String
    Dim tableData As Collection
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim colIndex As Integer

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

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

    ' Define the range for data in column A (A2:A14)
    Set dataRange = chartSheet.Range("A2:A14") ' Range for data (column A)

    ' Create a collection to hold valid data (filter out "false" or "falskt")
    Set tableData = New Collection
    validRowIndex = 1

    ' Loop through each row in column A (A2:A14)
    For rowIndex = 1 To dataRange.Rows.count
        cellValue = Trim(dataRange.Cells(rowIndex, 1).value)

        ' Filter out "false" or "falskt" and add valid values to tableData
        If LCase(cellValue) <> "false" And LCase(cellValue) <> "falskt" Then
            tableData.Add cellValue
            validRowIndex = validRowIndex + 1
        End If
    Next rowIndex

    ' If no valid rows found, show a message and exit
    If tableData.count = 0 Then
        MsgBox "No valid data to display.", vbExclamation
        Exit Sub
    End If

    ' Create a table on the slide with the filtered data (with "Yes" in column 2)
    pptSlide.Shapes.AddTable tableData.count, 2 ' Add rows based on filtered data, 2 columns
    Set tableShape = pptSlide.Shapes(pptSlide.Shapes.count)

    ' Name the table "Brands"
    tableShape.Name = "Brands"

    ' Format the table
    With tableShape.table
        ' Set column widths
        .Columns(1).width = 7 * 28.35 ' 7 cm (1 cm = 28.35 points)
        .Columns(2).width = 5 * 28.35 ' 5 cm

        ' Populate the table with filtered data and format each cell
        For rowIndex = 1 To tableData.count
            .cell(rowIndex, 1).shape.TextFrame.textRange.text = tableData.Item(rowIndex)
            .cell(rowIndex, 2).shape.TextFrame.textRange.text = "Yes" ' Default "Yes"

            For colIndex = 1 To 2
                ' Remove bold and set font color
                With .cell(rowIndex, colIndex).shape.TextFrame.textRange
                    .Font.Bold = msoFalse
                    .Font.color = RGB(17, 21, 66) ' Font color: RGB(17, 21, 66)
                End With

                ' Set fill color and border for each cell
                With .cell(rowIndex, colIndex).shape.Fill
                    .ForeColor.RGB = RGB(231, 232, 237) ' Fill color: #E7E8ED
                    .Solid
                End With

                ' Set border color and weight
                With .cell(rowIndex, colIndex).Borders
                    .Item(ppBorderTop).Weight = 0.25
                    .Item(ppBorderTop).ForeColor.RGB = RGB(17, 21, 66)
                    .Item(ppBorderBottom).Weight = 0.25
                    .Item(ppBorderBottom).ForeColor.RGB = RGB(17, 21, 66)
                    .Item(ppBorderLeft).Weight = 0.25
                    .Item(ppBorderLeft).ForeColor.RGB = RGB(17, 21, 66)
                    .Item(ppBorderRight).Weight = 0.25
                    .Item(ppBorderRight).ForeColor.RGB = RGB(17, 21, 66)
                End With
            Next colIndex
        Next rowIndex
    End With

    
End Sub

