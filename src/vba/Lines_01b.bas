Attribute VB_Name = "Lines_01b"
Sub Lines_1b()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim chartRow As Integer
    Dim operatingSystem As String
    Dim userName As String
    Dim FoundChart As Boolean
    Dim dataArray As Variant
    Dim i As Integer, j As Integer
    
    ' Get the current user's name and OS to set file path
    userName = Environ("USER")
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/line_chart_data_csv.csv"
    Else
        filePath = "C:\Local\line_chart_data_csv.csv"
    End If

    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the CSV file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Get the current slide and check for the second chart
    Set pptSlide = ActiveWindow.View.slide
    FoundChart = False

    Dim chartIndex As Integer
    chartIndex = 1 ' Set the chartIndex to 1 to start from the second chart

    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            If chartIndex = 2 Then ' Look for the second chart
                ' Get the workbook containing the chart's data
                Set chartDataWorkbook = chartShape.chart.chartData.Workbook
                Set chartSheet = chartDataWorkbook.Sheets(1)

                ' Clear the content starting from column U (preserve A to J)
                chartSheet.Range("U1:AV60").ClearContents ' Clear 60 rows from U to AV (21 columns)

                ' Read all data into an array
                rowIndex = 0
                chartRow = 1
                ReDim dataArray(1 To 60, 1 To 21) ' Store 60 rows of data for 21 columns

                Do While Not EOF(fileNumber) And rowIndex < 60 ' Limit to 60 rows
                    Line Input #fileNumber, line
                    rowIndex = rowIndex + 1

                    ' Split the line into data
                    Data = Split(line, ";")

                    ' Store data in the array
                    For colIndex = 1 To 21 ' First 21 columns (U to AV)
                        If colIndex <= UBound(Data) + 1 Then
                            ' Clean up unwanted characters
                            cellValue = Trim(Data(colIndex - 1))
                            If right(cellValue, 1) = "_" Or right(cellValue, 1) = "?" Then
                                cellValue = left(cellValue, Len(cellValue) - 1)
                            End If

                            dataArray(chartRow, colIndex) = cellValue
                        End If
                    Next colIndex
                    chartRow = chartRow + 1
                Loop

                ' Paste data into the Excel sheet in one go
                chartSheet.Range("U1").Resize(60, 21).value = dataArray

                ' Close the workbook
                chartShape.chart.chartData.Workbook.Close
                FoundChart = True
                Exit For
            Else
                chartIndex = chartIndex + 1
            End If
        End If
    Next chartShape

    If Not FoundChart Then
        MsgBox "No second chart found on the current slide."
    End If

    ' Close the CSV file
    Close fileNumber
End Sub

