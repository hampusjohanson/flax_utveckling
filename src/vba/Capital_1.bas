Attribute VB_Name = "Capital_1"
Sub Capital_111()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim startRow As Integer, endRow As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim chartRow As Integer
    Dim xValues As Variant
    Dim yValues As Variant
    Dim formulaRange As Object

    ' Debug print: Start execution
    Debug.Print "--- Starting Execution ---"

    ' Dynamically get the file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:/Local/exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    Debug.Print "File Path: " & filePath

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Debug.Print "Error: File not found"
        Exit Sub
    End If

    ' Set import range
    startRow = 21
    endRow = 40

    ' Open the file to read data
    fileNumber = FreeFile
    Debug.Print "Opening file: " & filePath
    Open filePath For Input As fileNumber

    ' Get the chart (embedded Excel) from the current slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Active Slide: " & pptSlide.SlideIndex

    On Error Resume Next
    For Each shape In pptSlide.Shapes
        Debug.Print "Checking shape: " & shape.Name
        If shape.hasChart Then
            Set chartShape = shape
            Debug.Print "Chart found: " & shape.Name
            Exit For
        End If
    Next shape
    On Error GoTo 0

    If chartShape Is Nothing Then
        MsgBox "No chart found on the current slide.", vbCritical
        Debug.Print "Error: No chart found"
        Exit Sub
    End If

    ' Open the chart's Excel workbook
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)
    Debug.Print "Opened chart data workbook"

    ' Clear previous data in range A2:C13
    Debug.Print "Clearing data range: A2:C13"
    chartSheet.Range("A2:C13").ClearContents

    ' Remove all existing series from the chart
    With chartShape.chart
        Do While .SeriesCollection.count > 0
            Debug.Print "Deleting series: " & .SeriesCollection(1).Name
            .SeriesCollection(1).Delete
        Loop
    End With

    ' Read data from CSV and paste into chart's data source
    rowIndex = 0
    chartRow = 2 ' Start from row 2 in Excel (A2:C2)

    Do While Not EOF(fileNumber) And chartRow <= 13
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' Debug: Print current line
        Debug.Print "Reading line " & rowIndex & ": " & line

        ' Skip rows outside the range
        If rowIndex < startRow Then GoTo SkipLine
        If rowIndex > endRow Then Exit Do

        ' Check if the row is empty or invalid
        If Trim(line) = "" Or line Like ";*" Then GoTo SkipLine

        ' Split the line into columns
        Data = Split(line, ";")

        ' Import only valid rows
        If rowIndex >= startRow And rowIndex <= endRow Then
            ' Skip rows where the first column is "false" or "falskt"
            If LCase(Trim(Data(0))) <> "false" And LCase(Trim(Data(0))) <> "falskt" And Trim(Data(0)) <> "" Then
                ' Insert Column 1 ? A
                chartSheet.Cells(chartRow, 1).value = Trim(Data(0))
                ' Insert Column 7 ? B (Convert to number)
                If IsNumeric(Trim(Data(6))) Then
                    chartSheet.Cells(chartRow, 2).Value2 = CDbl(Trim(Data(6)))
                Else
                    chartSheet.Cells(chartRow, 2).value = "NA"
                End If
                ' Insert Column 8 ? C (Convert to number)
                If IsNumeric(Trim(Data(7))) Then
                    chartSheet.Cells(chartRow, 3).Value2 = CDbl(Trim(Data(7)))
                Else
                    chartSheet.Cells(chartRow, 3).value = "NA"
                End If
                Debug.Print "Row " & chartRow & " A: " & Data(0) & " B: " & Data(6) & " C: " & Data(7)
                chartRow = chartRow + 1
            End If
        End If

SkipLine:
    Loop

    ' Close the file
    Close fileNumber
    Debug.Print "Finished reading file. Last row: " & chartRow

    ' Set new series using B2:B13 and C2:C13
    xValues = chartSheet.Range("B2:B13").value
    yValues = chartSheet.Range("C2:C13").value

    With chartShape.chart
        .SeriesCollection.newSeries
        .SeriesCollection(1).xValues = xValues
        .SeriesCollection(1).values = yValues

        ' Set markers to be round, size 14, and remove borders
        With .SeriesCollection(1)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 14
            .Format.line.visible = msoFalse
        End With
    End With

    ' Close the chart's Excel workbook
    chartShape.chart.chartData.Workbook.Close
    Debug.Print "Chart data workbook closed"

    ' Debug print: End execution
    Debug.Print "--- Execution Completed ---"

End Sub
