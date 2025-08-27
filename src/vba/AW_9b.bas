Attribute VB_Name = "AW_9b"
Sub ResizeAndRepositionChart()
    Dim sld As slide
    Dim shp As shape

    Set sld = ActiveWindow.View.slide

    ' Hitta första diagrammet på sliden
    For Each shp In sld.Shapes
        If shp.hasChart Then
            With shp
                .left = 2.49 * 28.35     ' cm till punkter
                .Top = 4 * 28.35
                .width = 18.64 * 28.35
                .height = 13.1 * 28.35
            End With
            Exit Sub
        End If
    Next shp
End Sub

