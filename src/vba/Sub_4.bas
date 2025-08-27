Attribute VB_Name = "Sub_4"
Sub Capital_1()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim chartRow As Integer
    Dim xValues As Variant
    Dim yValues As Variant
    Dim formulaRange As Object ' Change Range to Object to avoid error
    Dim operatingSystem As String
    Dim userName As String

    ' === Get the username from the Environ function ===
    userName = Environ("USER")
    
    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv" ' Use the username for macOS
    Else
        ' For Windows
        filePath = "C:\Users\" & userName & "\Desktop\exported_data_semi.csv" ' Build the file path for Windows
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Import specified part of CSV to the chart's Excel data ===
    startRow = 42 ' Start from row
    endRow = 91   ' End at row
    startCol = 1  ' Start from column
    endCol = 3    ' End at column

    ' Open the file to read data
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' === Get the chart (embedded Excel) from the current slide ===
    Set pptSlide = ActiveWindow.View.slide
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    On Error GoTo 0

    If chartShape Is Nothing Then
        MsgBox "No chart found on the current slide.", vbCritical
        Exit Sub
    End If

    ' === Open the chart's Excel workbook ===
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)

    ' === Clear the previous data range (A2:F51) in the chart's data source ===
    chartSheet.Range("A2:F51").ClearContents

    ' === Read data from CSV and paste into chart's data source ===
    rowIndex = 0
    chartRow = 2 ' Start from row 2 in Excel (row 1 is for headers)
    
    ' Read and fill data from CSV file
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1
        
        ' If the row is within the specified range
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            If LCase(Trim(Data(0))) <> "false" And Trim(Data(0)) <> "" Then
                ' Ignore empty rows and rows that start with "false"
                For colIndex = startCol To endCol
                    If colIndex <= UBound(Data) + 1 Then
                        cellValue = Trim(Data(colIndex - 1))
                        ' Remove "_" or "?" from the end of the text
                        If right(cellValue, 1) = "_" Or right(cellValue, 1) = "?" Then
                            cellValue = left(cellValue, Len(cellValue) - 1)
                        End If
                        chartSheet.Cells(chartRow, colIndex).value = cellValue
                    End If
                Next colIndex
                chartRow = chartRow + 1
            End If
        End If
    Loop

    ' Close the file
    Close fileNumber
    
    ' === Write the functions into column I ===
    Set formulaRange = chartSheet.Range("I2:I51") ' Adjust as necessary
    formulaRange.formula = "=IF(A2<>"""", B2+C2, """")" ' Example formula
    
    ' === Set new series using B2:B51 and C2:C51 ===
    xValues = chartSheet.Range("B2:B51").value
    yValues = chartSheet.Range("C2:C51").value

    With chartShape.chart
        .SeriesCollection.newSeries
        .SeriesCollection(1).xValues = xValues ' X-axis (B2:B51)
        .SeriesCollection(1).values = yValues ' Y-axis (C2:C51)

        ' Set markers to be round, size 14, and remove borders
        With .SeriesCollection(1)
            .MarkerStyle = xlMarkerStyleCircle ' Round markers
            .MarkerSize = 5 ' Size 14
            .Format.line.visible = msoFalse ' Remove the borders from markers
        End With
    End With

    ' Close the chart's Excel workbook
    chartShape.chart.chartData.Workbook.Close

End Sub

