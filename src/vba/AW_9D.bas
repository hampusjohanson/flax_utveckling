Attribute VB_Name = "AW_9D"
Sub CreateBrandTableOnSlide()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim brandTable As shape
    Dim wb As Object
    Dim ws As Object
    Dim i As Integer, j As Integer, k As Integer
    Dim val As Variant
    Dim brandList As Collection
    Dim cell As cell
    Dim borderColor As Long
    Dim bgColor As Long
    Dim borderTypes(1 To 4) As PpBorderType

    Set sld = ActiveWindow.View.slide

    ' Hämta första diagrammet
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set cht = shp.chart
            Exit For
        End If
    Next shp
    If cht Is Nothing Then
        MsgBox "Inget diagram hittades.", vbExclamation
        Exit Sub
    End If

    ' Hämta Excel-data
    Set wb = cht.chartData.Workbook
    Set ws = wb.Sheets(1)

    Set brandList = New Collection
    For i = 3 To 18 ' AA3:AA18
        val = ws.Range("AA" & i).value
        If VarType(val) = vbString Or VarType(val) = vbVariant Then
            If LCase(Trim(val)) <> "false" And LCase(Trim(val)) <> "falskt" And Trim(val) <> "" Then
                brandList.Add val
            End If
        End If
    Next i

    If brandList.count = 0 Then
        MsgBox "Inga giltiga varumärken hittades i AA3:AA18.", vbExclamation
        Exit Sub
    End If

    ' Skapa tabell på sliden
    Set brandTable = sld.Shapes.AddTable(brandList.count, 1)
    brandTable.Name = "BRANDS_SEL"

    With brandTable
        .left = 50
        .Top = 100
        .width = 150
        .height = brandList.count * 20
    End With

    ' Ta bort all tabellformat (t.ex. header row styles)
    brandTable.table.ApplyStyle "", True

    ' Färger
    borderColor = RGB(16, 21, 66)
    bgColor = RGB(240, 240, 240)

    borderTypes(1) = ppBorderTop
    borderTypes(2) = ppBorderBottom
    borderTypes(3) = ppBorderLeft
    borderTypes(4) = ppBorderRight

    ' Formatera och fyll celler
    For j = 1 To brandList.count
        Set cell = brandTable.table.cell(j, 1)

        ' Text
        With cell.shape.TextFrame.textRange
            .text = brandList(j)
            With .Font
                .size = 14
                .Bold = msoFalse
                .color.RGB = borderColor
            End With
        End With

        ' Bakgrund
        cell.shape.Fill.ForeColor.RGB = bgColor

        ' Ta bort och lägg till kantlinjer
        For k = 1 To 4
            With cell.Borders(borderTypes(k))
                .visible = msoTrue
                .ForeColor.RGB = borderColor
                .Weight = 1
            End With
        Next k
    Next j
End Sub

