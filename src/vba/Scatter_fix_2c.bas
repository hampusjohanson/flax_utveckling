Attribute VB_Name = "Scatter_fix_2c"
' === Modul: Scatter_fix_2c ===

Option Explicit

Sub Scatter_fix_2c()

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim ser As Object
    Dim pt As Object
    Dim i As Integer
    Dim s As Integer
    Dim chartFound As Boolean

    Dim tableShape As shape
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim rowIndex As Integer

    Dim labelText As String
    Dim pointLeft As Single
    Dim pointTop As Single

    Set sld = ActiveWindow.View.slide
    chartFound = False

    ' Leta upp kopian
    For Each shp In sld.Shapes
        If shp.Type = msoChart Then
            If shp.Name = "kopia_chart" Then
                Set cht = shp.chart
                chartFound = True
                Exit For
            End If
        End If
    Next shp

    If Not chartFound Then
        Debug.Print "Chart 'kopia_chart' not found on this slide."
        Exit Sub
    End If

    ' Skapa ny tabell och namnge den
    rowCount = 51
    colCount = 4
    Set tableShape = sld.Shapes.AddTable(rowCount, colCount, 800, 50, 400, 400)
    tableShape.Name = "points_table"

    ' Rubriker
    With tableShape.table
        .cell(1, 1).shape.TextFrame.textRange.text = "#"
        .cell(1, 2).shape.TextFrame.textRange.text = "Text"
        .cell(1, 3).shape.TextFrame.textRange.text = "PointLeft"
        .cell(1, 4).shape.TextFrame.textRange.text = "PointTop"
    End With

    ' Radnummer
    For rowIndex = 2 To rowCount
        tableShape.table.cell(rowIndex, 1).shape.TextFrame.textRange.text = CStr(rowIndex - 1)
    Next rowIndex

    ' Läs punkterna med korrekt loop
    For s = 1 To cht.SeriesCollection.count
        Set ser = cht.SeriesCollection(s)
        
        For i = 1 To ser.Points.count
            Set pt = ser.Points(i)

            On Error Resume Next
            If pt.HasDataLabel Then
                labelText = pt.dataLabel.text
                
                ' Hämta datapunktens vänster och topp från punktens koordinater i diagrammets pixelområde
                pointLeft = pt.left
                pointTop = pt.Top
                
                ' Fyll tabellen
                tableShape.table.cell(i + 1, 2).shape.TextFrame.textRange.text = labelText
                tableShape.table.cell(i + 1, 3).shape.TextFrame.textRange.text = Format(pointLeft, "0.00")
                tableShape.table.cell(i + 1, 4).shape.TextFrame.textRange.text = Format(pointTop, "0.00")
            End If
            On Error GoTo 0

        Next i
    Next s

    Debug.Print "Finished creating point-position table."

End Sub

