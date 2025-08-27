Attribute VB_Name = "Capital_8_Add_trend"
Sub Mac_Cap_trendline()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim trendline As trendline

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Hitta diagrammet på sliden
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Ta bort befintliga trendlinjer
    With chartShape.chart.SeriesCollection(1)
        Do While .Trendlines.count > 0
            .Trendlines(1).Delete
        Loop
    End With

    ' Lägg till en trendlinje
    With chartShape.chart.SeriesCollection(1).Trendlines.Add(Type:=xlLinear)
        .Format.line.Weight = 0.25 ' Linjetjocklek
        .Format.line.ForeColor.RGB = RGB(17, 21, 66) ' Färg
        .Format.line.DashStyle = msoLineLongDash ' Lång streckad linje
        .Forward = 1 ' Prognos framåt (använd 1 för att säkerställa att det fungerar)
        .Backward = 1 ' Prognos bakåt (använd 1 för att säkerställa att det fungerar)
    End With

End Sub

