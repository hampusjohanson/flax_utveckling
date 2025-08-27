Attribute VB_Name = "Mac_SP_2"

Sub Mac_SP_2()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim sourceTable As table
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim sheetName As String
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer
    Dim dataCount As Integer
    Dim labels() As String, xData() As Double, yData() As Double, bubbleSize() As Double, bubbleColors() As String
    Dim isBubbleChart As Boolean

    ' === Steg 1: Importera specifik del av CSV som en ny tabell ===
    filePath = "/Users/hampus.johansson/Desktop/exported_data_semi.csv"

    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    startRow = 1
    endRow = 15
    startCol = 1
    endCol = 6

    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable(numRows:=endRow - startRow + 1, NumColumns:=endCol - startCol + 1, left:=50, Top:=50, width:=600, height:=300)
    Set sourceTable = importedTableShape.table

    rowIndex = 0
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    sourceTable.cell(rowIndex - startRow + 1, colIndex - startCol + 1).shape.TextFrame.textRange.text = Trim(Data(colIndex - 1))
                End If
            Next colIndex
        End If
    Loop
    Close fileNumber

    dataCount = 0
    For rowIndex = 1 To sourceTable.Rows.count
        If LCase(Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) = "false" Then Exit For
        dataCount = dataCount + 1
    Next rowIndex

    ReDim labels(1 To dataCount)
    ReDim xData(1 To dataCount)
    ReDim yData(1 To dataCount)
    ReDim bubbleSize(1 To dataCount)
    ReDim bubbleColors(1 To dataCount)

    For rowIndex = 1 To dataCount
        labels(rowIndex) = Replace(Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text), "_", "") ' Ta bort "_"
        xData(rowIndex) = val(sourceTable.cell(rowIndex, 3).shape.TextFrame.textRange.text) / 100
        yData(rowIndex) = val(sourceTable.cell(rowIndex, 2).shape.TextFrame.textRange.text) / 100
        bubbleSize(rowIndex) = val(sourceTable.cell(rowIndex, 5).shape.TextFrame.textRange.text)
        bubbleColors(rowIndex) = Trim(sourceTable.cell(rowIndex, 6).shape.TextFrame.textRange.text)
    Next rowIndex

    ' === Steg 2: Hitta befintligt diagram ===
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

    ' Kontrollera om det Šr ett bubble chart
    isBubbleChart = chartShape.chart.chartType = xlBubble

    ' …ppna diagrammets data
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Worksheets(1)
    sheetName = chartSheet.Name

    chartSheet.Cells.Clear
    chartSheet.Cells(1, 1).value = "Labels"
    chartSheet.Cells(1, 2).value = "X Values"
    chartSheet.Cells(1, 3).value = "Y Values"
    If isBubbleChart Then chartSheet.Cells(1, 4).value = "Bubble Size"

    For rowIndex = 1 To dataCount
        chartSheet.Cells(rowIndex + 1, 1).value = labels(rowIndex)
        chartSheet.Cells(rowIndex + 1, 2).value = xData(rowIndex)
        chartSheet.Cells(rowIndex + 1, 3).value = yData(rowIndex)
        If isBubbleChart Then chartSheet.Cells(rowIndex + 1, 4).value = bubbleSize(rowIndex)
    Next rowIndex

    With chartShape.chart
        .Axes(xlCategory).MinimumScale = 0
        .Axes(xlCategory).MaximumScale = 1
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1
        .SeriesCollection(1).xValues = "='" & sheetName & "'!B2:B" & (dataCount + 1)
        .SeriesCollection(1).values = "='" & sheetName & "'!C2:C" & (dataCount + 1)
        If isBubbleChart Then
            .SeriesCollection(1).BubbleSizes = "='" & sheetName & "'!D2:D" & (dataCount + 1)
        Else
            .SeriesCollection(1).Format.line.visible = msoFalse ' Ta bort linjefŠrg fšr scatter
        End If
    End With

    chartDataWorkbook.Close

    ' SŠtt etiketter och fŠrger
    With chartShape.chart
        .SeriesCollection(1).ApplyDataLabels
        For rowIndex = 1 To dataCount
            If rowIndex <= .SeriesCollection(1).Points.count Then
                With .SeriesCollection(1).Points(rowIndex)
                    .dataLabel.text = labels(rowIndex)
                    .Format.line.visible = msoFalse ' Ta bort kontur runt bubblan/scatter
                    Dim hexColor As String
                    hexColor = bubbleColors(rowIndex)
                    If Len(hexColor) = 7 And left(hexColor, 1) = "#" Then
                        Dim r As Long, g As Long, b As Long
                        r = CLng("&H" & Mid(hexColor, 2, 2))
                        g = CLng("&H" & Mid(hexColor, 4, 2))
                        b = CLng("&H" & Mid(hexColor, 6, 2))
                        .Format.Fill.ForeColor.RGB = RGB(r, g, b)
                    End If
                End With
            End If
        Next rowIndex
    End With

    ' Ta bort outlines frŒn alla datapunkter
    With chartShape.chart
        For Each series In .SeriesCollection
            For Each Point In series.Points
                Point.Format.line.ForeColor.RGB = RGB(255, 255, 255) ' SŠtt linjefŠrgen till vitt som en workaround
                Point.Format.line.Transparency = 1 ' Gšr linjen helt transparent
                Point.Format.line.visible = msoFalse ' Dšlja linjen helt
            Next Point
        Next series
    End With

    importedTableShape.Delete

End Sub
