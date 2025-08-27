Attribute VB_Name = "AW_A3"
Sub AW_A3()
    On Error Resume Next ' Tyst ignorera alla fel

    Dim pptSlide As slide
    Dim tableShape As shape
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim dataArray As Variant
    Dim sourceRowIndex As Integer, sourceColIndex As Integer

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = pptSlide.Shapes("SOURCE")
    If tableShape Is Nothing Then Exit Sub

    ' Hitta diagrammet och koppla till Excel
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartDataWorkbook = chartShape.chart.chartData.Workbook
            Set chartSheet = chartDataWorkbook.Sheets(1)
            Exit For
        End If
    Next chartShape
    If chartSheet Is Nothing Then Exit Sub

    ' Kopiera data från tabellen till array
    ReDim dataArray(1 To 22, 1 To 7)
    For sourceRowIndex = 1 To 22
        For sourceColIndex = 1 To 7
            cellValue = tableShape.table.cell(sourceRowIndex, sourceColIndex).shape.TextFrame.textRange.text

            If InStr(cellValue, "%") > 0 Then
                cellValue = Replace(cellValue, "%", "")
                If IsNumeric(cellValue) Then
                    dataArray(sourceRowIndex, sourceColIndex) = CDbl(cellValue) / 100
                Else
                    dataArray(sourceRowIndex, sourceColIndex) = cellValue
                End If
            ElseIf IsNumeric(cellValue) Then
                dataArray(sourceRowIndex, sourceColIndex) = CDbl(cellValue)
            Else
                dataArray(sourceRowIndex, sourceColIndex) = cellValue
            End If
        Next sourceColIndex
    Next sourceRowIndex

    ' Klistra in i Excel-arket
    With chartSheet.Range("AA1").Resize(22, 7)
        .value = dataArray
        .NumberFormat = "General"
    End With

    ' Ta bort tabellen
    tableShape.Delete

End Sub

