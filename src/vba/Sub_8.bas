Attribute VB_Name = "Sub_8"
Sub Sub_8()
    Dim pptSlide As slide
    Dim userName As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim sourceTable As table
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim tableRow As Integer, tableCol As Integer
    Dim cellValue As String
    Dim operatingSystem As String
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim shape As shape

    ' Determine operating system and set file path
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        Set pptSlide = ActivePresentation.Slides(1)
        If pptSlide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
            userName = Trim(Split(pptSlide.NotesPage.Shapes.Placeholders(2).TextFrame.textRange.text, vbCrLf)(0))
        Else
            userName = Environ$("USER")
        End If
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Define hardcoded rows and columns
    startRow = 342
    endRow = 391
    startCol = 1
    endCol = 11

    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Add a new table to the slide
    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable(numRows:=endRow - startRow + 1, NumColumns:=endCol - startCol + 1, left:=50, Top:=50, width:=600, height:=300)
    importedTableShape.Name = "table_substance"
    Set sourceTable = importedTableShape.table

    ' Read data and fill the table
    tableRow = 1
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        If tableRow >= startRow And tableRow <= endRow Then
            Data = Split(line, ";")
            For tableCol = startCol To endCol
                If tableCol <= UBound(Data) + 1 Then
                    cellValue = Trim(Data(tableCol - 1))
                    If IsNumeric(cellValue) Then cellValue = Format(CDbl(cellValue), "0.000")
                    sourceTable.cell(tableRow - startRow + 1, tableCol).shape.TextFrame.textRange.text = cellValue
                End If
            Next tableCol
        End If
        tableRow = tableRow + 1
    Loop
    Close fileNumber

    ' Clean rows with "false" in the first column
    For rowIndex = sourceTable.Rows.count To 1 Step -1
        If LCase(Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) = "false" Then
            sourceTable.Rows(rowIndex).Delete
        End If
    Next rowIndex

    ' Remove trailing "_" and "?" in the first column
    For rowIndex = 1 To sourceTable.Rows.count
        cellValue = sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text
        If right(cellValue, 1) = "_" Or right(cellValue, 1) = "?" Then
            sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text = left(cellValue, Len(cellValue) - 1)
        End If
    Next rowIndex

    ' Replace "false" with empty text in all cells
    For rowIndex = 1 To sourceTable.Rows.count
        For colIndex = 1 To sourceTable.Columns.count
            If LCase(Trim(sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text)) = "false" Then
                sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = ""
            End If
        Next colIndex
    Next rowIndex

    ' Find chart on the slide
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape

    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Activate chart data and clean Excel range
    On Error Resume Next
    chartShape.chart.chartData.Activate
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo 0

    If chartDataWorkbook Is Nothing Then
        MsgBox "Unable to open chart data. Check compatibility with macOS.", vbCritical
        Exit Sub
    End If

    Set chartSheet = chartDataWorkbook.Worksheets(1)
    chartSheet.Range("J2:Q52").Clear

    ' Copy specific columns to Excel
    For rowIndex = 1 To sourceTable.Rows.count
        chartSheet.Cells(rowIndex + 1, 10).value = sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text
        chartSheet.Cells(rowIndex + 1, 12).value = sourceTable.cell(rowIndex, 2).shape.TextFrame.textRange.text
        chartSheet.Cells(rowIndex + 1, 11).value = sourceTable.cell(rowIndex, 5).shape.TextFrame.textRange.text
        For colIndex = 6 To 10
            chartSheet.Cells(rowIndex + 1, colIndex + 7).value = sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text
        Next colIndex
    Next rowIndex

    ' Remove the temporary table from the slide
    importedTableShape.Delete

    ' Force chart to visually update
    On Error Resume Next
    chartShape.chart.chartData.Workbook.Close False ' Close the embedded data window
    chartShape.chart.chartData.Activate ' Reactivate chart data to refresh visuals
    chartShape.chart.chartData.Workbook.Close False
    On Error GoTo 0

 End Sub

