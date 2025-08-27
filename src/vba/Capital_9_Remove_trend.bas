Attribute VB_Name = "Capital_9_Remove_trend"
Sub Mac_Cap_Remove_Trendline()
    Dim pptSlide As slide
    Dim chartShape As shape

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Hitta diagrammet p� sliden
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades p� sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Ta bort alla trendlinjer i det f�rsta dataserien i diagrammet
    With chartShape.chart.SeriesCollection(1)
        Do While .Trendlines.count > 0
            .Trendlines(1).Delete
        Loop
    End With

End Sub

