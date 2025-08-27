Attribute VB_Name = "LD_2"
Sub LD_2()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim horizontalCrossing As Double, verticalCrossing As Double
    Dim rowIndex As Integer

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

    ' Hämta diagrammets datakälla
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1)

    ' Definiera ett namngivet område för A1:F15
    On Error Resume Next
    chartSheet.Names.Add Name:="DataRange", RefersTo:=chartSheet.Range("A1:F15")
    On Error GoTo 0

    ' Skriv formler i celler i kolumn I
    With chartSheet
        .Range("I2").formula = "=MIN(B2:B51)-I9"
        .Range("I3").formula = "=MAX(B2:B51)+I9"
        .Range("I4").formula = "=MIN(C2:C51)-I9"
        .Range("I5").formula = "=MAX(C2:C51)+I9"
        .Range("I6").formula = "=MEDIAN(B2:B51)"
        .Range("I7").formula = "=MEDIAN(C2:C51)"
        .Range("I9").value = 0.02
    End With

    ' Läs axelvärden från kolumn I
    xMin = chartSheet.Range("I2").value
    xMax = chartSheet.Range("I3").value
    yMin = chartSheet.Range("I4").value
    yMax = chartSheet.Range("I5").value
    horizontalCrossing = chartSheet.Range("I6").value
    verticalCrossing = chartSheet.Range("I7").value

    ' Justera axlar i diagrammet
    With chartShape.chart
        ' Ställ in skalor och korsningspunkter
        .Axes(xlCategory).MinimumScale = xMin
        .Axes(xlCategory).MaximumScale = xMax
        .Axes(xlCategory).CrossesAt = horizontalCrossing
        .Axes(xlValue).MinimumScale = yMin
        .Axes(xlValue).MaximumScale = yMax
        .Axes(xlValue).CrossesAt = verticalCrossing

        ' Anpassa utseendet på axellinjerna
        With .Axes(xlCategory).Format.line
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66) ' Färg
            .DashStyle = msoLineLongDash ' Lång streckad linje
            .Weight = 0.25 ' Linjetjocklek
        End With

        With .Axes(xlValue).Format.line
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66) ' Färg
            .DashStyle = msoLineLongDash ' Lång streckad linje
            .Weight = 0.25 ' Linjetjocklek
        End With
    End With

    ' Stäng diagrammets datakälla
    chartShape.chart.chartData.Workbook.Close

End Sub



