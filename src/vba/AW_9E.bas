Attribute VB_Name = "AW_9E"
Sub Paste_BRANDS_SEL_to_Excel()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim wb As Object
    Dim ws As Object
    Dim pptTable As table
    Dim tblShape As shape
    Dim i As Integer
    Dim found As Boolean

    Set sld = ActiveWindow.View.slide
    found = False

    ' Hitta tabellen BRANDS_SEL
    For Each shp In sld.Shapes
        If shp.HasTable Then
            If shp.Name = "BRANDS_SEL" Then
                Set pptTable = shp.table
                Set tblShape = shp
                found = True
                Exit For
            End If
        End If
    Next shp

    If Not found Then
        MsgBox "Tabellen 'BRANDS_SEL' hittades inte på sliden.", vbExclamation
        Exit Sub
    End If

    ' Hitta första diagrammet på sliden
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set cht = shp.chart
            Exit For
        End If
    Next shp

    If cht Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbExclamation
        Exit Sub
    End If

    Set wb = cht.chartData.Workbook
    Set ws = wb.Worksheets(1)

    ' ?? Rensa tidigare innehåll i A35:A50
    ws.Range("A35:A50").ClearContents

    ' ?? Klistra in nya varumärken från BRANDS_SEL
    For i = 1 To pptTable.Rows.count
        ws.Cells(34 + i, 1).value = pptTable.cell(i, 1).shape.TextFrame.textRange.text
    Next i

    ' ?? Kör uppdateringsmakrot
    Application.Run "Awareness_Get_Data"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' ??? Ta bort tabellen från sliden
    tblShape.Delete
End Sub

