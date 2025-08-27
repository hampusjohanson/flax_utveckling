Attribute VB_Name = "AW_9"
Sub IncreaseTopSeriesByHalfPercent()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim srs As series
    Dim vals As Variant
    Dim i As Integer

    Set sld = ActiveWindow.View.slide

    ' Hitta första diagrammet på sliden
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set cht = shp.chart
            Exit For
        End If
    Next shp

    If cht Is Nothing Then Exit Sub

    ' Välj sista serien i diagrammet (översta visuellt)
    With cht.SeriesCollection(cht.SeriesCollection.count)
        vals = .values
        For i = LBound(vals) To UBound(vals)
            If IsNumeric(vals(i)) Then
                vals(i) = vals(i) + 0.01
            Else
                vals(i) = 0.01
            End If
        Next i
        .values = vals
    End With
End Sub

Sub DecreaseTopSeriesByOnePercent()
    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim vals As Variant
    Dim i As Integer

    Set sld = ActiveWindow.View.slide

    ' Hitta första diagrammet på sliden
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set cht = shp.chart
            Exit For
        End If
    Next shp

    If cht Is Nothing Then Exit Sub

    ' Välj sista serien i diagrammet (översta visuellt)
    With cht.SeriesCollection(cht.SeriesCollection.count)
        vals = .values
        For i = LBound(vals) To UBound(vals)
            If IsNumeric(vals(i)) Then
                vals(i) = vals(i) - 0.01
                If vals(i) < 0 Then vals(i) = 0
            Else
                vals(i) = 0
            End If
        Next i
        .values = vals
    End With
End Sub

