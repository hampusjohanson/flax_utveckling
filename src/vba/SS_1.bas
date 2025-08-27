Attribute VB_Name = "SS_1"
Sub SS_1()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim shape As shape
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
    Dim operatingSystem As String
    Dim userName As String
    Dim skippedRows As Integer
    Dim lastRow As Long

    skippedRows = 0

    ' === Get the username from the Environ function ===
    userName = Environ("USER")
    
    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    startRow = 472
    endRow = 521
    startCol = 1
    endCol = 3

    fileNumber = FreeFile
    Open filePath For Input As fileNumber

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
        Close fileNumber
        Exit Sub
    End If

    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)

    chartSheet.Range("K2:M51").ClearContents

    rowIndex = 0
    chartRow = 2

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1
        
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            If LCase(Trim(Data(0))) <> "false" And LCase(Trim(Data(0))) <> "falskt" And Trim(Data(0)) <> "" Then
                For colIndex = startCol To endCol
                    If colIndex <= UBound(Data) + 1 Then
                        cellValue = Trim(Data(colIndex - 1))
                        If right(cellValue, 1) = "_" Or right(cellValue, 1) = "?" Then
                            cellValue = left(cellValue, Len(cellValue) - 1)
                        End If
                        chartSheet.Cells(chartRow, colIndex + 10).value = cellValue ' Column K = 11
                    End If
                Next colIndex
                chartRow = chartRow + 1
            Else
                skippedRows = skippedRows + 1
            End If
        End If
    Loop

    Close fileNumber

    lastRow = chartRow - 1

    xValues = chartSheet.Range("K2:K" & lastRow).value
    yValues = chartSheet.Range("L2:L" & lastRow).value

    With chartShape.chart
        .SeriesCollection.newSeries
        With .SeriesCollection(1)
            .xValues = xValues
            .values = yValues
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .Format.line.visible = msoFalse
        End With
    End With

    chartDataWorkbook.Close

    Debug.Print "SS_1 färdig. Skippade rader på grund av 'false' eller tomma: " & skippedRows
    Debug.Print "Antal datapunkter: " & (lastRow - 1)
End Sub

