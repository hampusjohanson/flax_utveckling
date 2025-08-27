Attribute VB_Name = "AW_8"
Sub AddTransparentAndGreenOverlayFromExcel()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim ws As Object
    Dim i As Integer, j As Integer
    Dim val As Variant
    Dim numPoints As Integer
    Dim baseValues() As Double
    Dim overlayValues() As Double

    ' Hämta första diagrammet
    Set sld = ActiveWindow.View.slide
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
    Set ws = cht.chartData.Workbook.Sheets(1)
    numPoints = cht.SeriesCollection(1).Points.count
    ReDim baseValues(1 To numPoints)
    ReDim overlayValues(1 To numPoints)

    ' Rensa tidigare overlay-serier
    For j = cht.SeriesCollection.count To 1 Step -1
        Select Case cht.SeriesCollection(j).Name
            Case "Overlay", "Overlay_Green"
                cht.SeriesCollection(j).Delete
        End Select
    Next j

    ' Läs in värden från G (färglös bas)
    For i = 1 To numPoints
        val = ws.Range("G" & i + 1).value
        If IsNumeric(val) Then baseValues(i) = val Else baseValues(i) = 0
    Next i

    ' Lägg till färglös overlay
    With cht.SeriesCollection.newSeries
        .values = baseValues
        .Name = "Overlay"
        .chartType = xlColumnStacked
        .HasDataLabels = False
        .Format.Fill.visible = msoFalse
        .Format.line.visible = msoFalse
    End With

    ' Läs in värden från H (grön topp)
    For i = 1 To numPoints
        val = ws.Range("H" & i + 1).value
        If IsNumeric(val) Then overlayValues(i) = val Else overlayValues(i) = 0
    Next i

 ' Lägg till grön overlay
With cht.SeriesCollection.newSeries
    .values = overlayValues
    .Name = "Overlay_Green"
    .chartType = xlColumnStacked
    .HasDataLabels = True
    .Format.Fill.visible = msoTrue
    .Format.Fill.Solid
    .Format.Fill.ForeColor.RGB = RGB(136, 255, 194) ' #88FFC2
    .Format.line.visible = msoFalse
    
    ' Anpassa datalabels
    With .DataLabels
        .Font.Bold = True
        .Font.color = RGB(17, 21, 66)
    End With
End With

End Sub
