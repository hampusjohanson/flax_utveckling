Attribute VB_Name = "Scatter_fix_4"
' === Modul: Scatter_fix_4 ===

Option Explicit

Sub Scatter_fix_4()

    Dim sld As slide
    Dim shp As shape
    Dim tableShape As shape
    Dim tbl As table
    Dim chartShape As shape
    Dim cht As chart
    Dim wb As Object ' Excel.Workbook
    Dim ws As Object ' Excel.Worksheet
    
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim i As Integer, j As Integer

    Set sld = ActiveWindow.View.slide
    Set tableShape = Nothing
    Set chartShape = Nothing
    
    ' Leta upp labels_table på sliden
    For Each shp In sld.Shapes
        If shp.HasTable Then
            If shp.Name = "labels_table" Then
                Set tableShape = shp
                Exit For
            End If
        End If
    Next shp
    
    If tableShape Is Nothing Then
        Debug.Print "Tabellen 'labels_table' hittades inte på denna slide."
        Exit Sub
    End If
    
    Set tbl = tableShape.table
    rowCount = tbl.Rows.count
    colCount = tbl.Columns.count
    
    ' Leta upp kopia_chart på samma slide
    For Each shp In sld.Shapes
        If shp.Name = "kopia_excel_chart" Then
            Set chartShape = shp
            Exit For
        End If
    Next shp
    
    If chartShape Is Nothing Then
        Debug.Print "Diagrammet 'kopia_excel_chart' hittades inte på sliden."
        Exit Sub
    End If
    
    ' Hämta inbäddade arbetsboken
    Set cht = chartShape.chart
    cht.chartData.Activate ' Viktigt: aktivera datakopplingen
    Set wb = cht.chartData.Workbook
    Set ws = wb.Sheets(1) ' Oftast heter det redan Sheet1

    ' Rensa gamla data
    ws.Cells.Clear

    ' Klistra in tabellen från labels_table
    For i = 1 To rowCount
        For j = 1 To colCount
            ws.Cells(i, j).value = tbl.cell(i, j).shape.TextFrame.textRange.text
        Next j
    Next i
    
    Debug.Print "Tabellen 'labels_table' har nu infogats i diagrammets Excel-data."

End Sub

