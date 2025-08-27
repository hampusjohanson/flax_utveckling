Attribute VB_Name = "Capital_8_Add_trend"
Sub Mac_Cap_trendline()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim trendline As trendline

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

    ' Ta bort befintliga trendlinjer
    With chartShape.chart.SeriesCollection(1)
        Do While .Trendlines.count > 0
            .Trendlines(1).Delete
        Loop
    End With

    ' L�gg till en trendlinje
    With chartShape.chart.SeriesCollection(1).Trendlines.Add(Type:=xlLinear)
        .Format.line.Weight = 0.25 ' Linjetjocklek
        .Format.line.ForeColor.RGB = RGB(17, 21, 66) ' F�rg
        .Format.line.DashStyle = msoLineLongDash ' L�ng streckad linje
        .Forward = 1 ' Prognos fram�t (anv�nd 1 f�r att s�kerst�lla att det fungerar)
        .Backward = 1 ' Prognos bak�t (anv�nd 1 f�r att s�kerst�lla att det fungerar)
    End With

End Sub

