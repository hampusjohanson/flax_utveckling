Attribute VB_Name = "AW_9A"
Sub SetLabelToSumOfTwoBottomSeries()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim topSeries As series
    Dim bottomSeries1 As series
    Dim bottomSeries2 As series
    Dim i As Integer
    Dim vals1 As Variant, vals2 As Variant
    Dim valSum As Double

    Set sld = ActiveWindow.View.slide

    ' Hitta första diagrammet på sliden
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set cht = shp.chart
            Exit For
        End If
    Next shp

    If cht Is Nothing Then Exit Sub
    If cht.SeriesCollection.count < 3 Then Exit Sub

    Set bottomSeries1 = cht.SeriesCollection(1)
    Set bottomSeries2 = cht.SeriesCollection(2)
    Set topSeries = cht.SeriesCollection(cht.SeriesCollection.count)

    vals1 = bottomSeries1.values
    vals2 = bottomSeries2.values

    With topSeries
        .HasDataLabels = True
        For i = LBound(vals1) To UBound(vals1)
            If IsNumeric(vals1(i)) And IsNumeric(vals2(i)) Then
                valSum = vals1(i) + vals2(i)
                .DataLabels(i).text = Format(valSum, "0%")
                .DataLabels(i).Font.color = RGB(17, 21, 66)
            End If
        Next i
    End With
End Sub

