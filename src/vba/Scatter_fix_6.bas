Attribute VB_Name = "Scatter_fix_6"
' === Modul: Scatter_fix_6 ===

Option Explicit

Sub Scatter_fix_6()

    Dim sld As slide
    Dim shp As shape
    Dim chartShape As shape
    Dim tableShape As shape
    Dim tbl As table
    Dim cht As chart
    Dim wb As Object ' Excel.Workbook
    Dim ws As Object ' Excel.Worksheet
    
    Dim i As Integer
    Dim pointLeft As String
    Dim pointTop As String
    Dim rowCount As Integer

    Set sld = ActiveWindow.View.slide
    Set chartShape = Nothing
    Set tableShape = Nothing

    ' Hitta chart
    For Each shp In sld.Shapes
        If shp.Name = "kopia_excel_chart" Then
            Set chartShape = shp
            Exit For
        End If
    Next shp

    If chartShape Is Nothing Then
        Debug.Print "'kopia_excel_chart' hittades inte."
        Exit Sub
    End If

    ' Hitta points_table
    For Each shp In sld.Shapes
        If shp.Name = "points_table" And shp.HasTable Then
            Set tableShape = shp
            Exit For
        End If
    Next shp

    If tableShape Is Nothing Then
        Debug.Print "'points_table' hittades inte."
        Exit Sub
    End If

    Set tbl = tableShape.table
    rowCount = tbl.Rows.count

    ' Öppna Excel inbäddat i diagram
    Set cht = chartShape.chart
    cht.chartData.Activate
    Set wb = cht.chartData.Workbook
    Set ws = wb.Sheets(1)

    ' Klistra in kolumn 3 och 4 i kolumn I och J
    For i = 2 To rowCount ' börja på rad 2, skippar rubriker
        pointLeft = tbl.cell(i, 3).shape.TextFrame.textRange.text
        pointTop = tbl.cell(i, 4).shape.TextFrame.textRange.text

        ws.Cells(i, 9).value = pointLeft   ' kolumn I
        ws.Cells(i, 10).value = pointTop   ' kolumn J
    Next i

    ' Radera tabellen från sliden
    tableShape.Delete
    Debug.Print "'points_table' har tagits bort."

    ' Stäng kopplingen till inbäddade Excel-data
    wb.Application.Quit
    Debug.Print "Exceldata för 'kopia_excel_chart' stängd."

End Sub

