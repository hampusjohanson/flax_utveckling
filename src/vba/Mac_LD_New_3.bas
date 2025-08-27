Attribute VB_Name = "Mac_LD_New_3"
Sub LD_3()
    ' Justera axlar
    
    Dim pptSlide As slide
    Dim embeddedChart As shape
    Dim chartWorkbook As Object ' Excel arbetsboken för diagramdata
    Dim chartSheet As Object ' Referens till Excel-arket
    Dim shapeIndex As Integer
    Dim shapeType As Integer
    Dim xMin As Double, xMax As Double
    Dim yMin As Double, yMax As Double
    Dim horizontalCrossing As Double
    Dim verticalCrossing As Double
    Dim xMedian As Double
    Dim operatingSystem As String
    Dim topPosition As Single

    ' Hitta diagrammet på sliden
    Set pptSlide = ActiveWindow.View.slide

    ' Sök igenom alla objekt på sliden för att hitta ett diagram
    For shapeIndex = 1 To pptSlide.Shapes.count
        shapeType = pptSlide.Shapes(shapeIndex).Type
        If shapeType = msoChart Then
            Set embeddedChart = pptSlide.Shapes(shapeIndex)
            Exit For ' Hitta och avsluta loopen
        End If
    Next shapeIndex

    ' Om ett diagram finns, fortsätt
    If Not embeddedChart Is Nothing Then
        ' Hämta Excel arbetsboken och arket för diagrammet
        Set chartWorkbook = embeddedChart.chart.chartData.Workbook
        Set chartSheet = chartWorkbook.Sheets(1)

        ' Skapa formler i kolumn I för att beräkna min- och max-värden samt medianer
        With chartSheet
            .Range("I2").formula = "=MIN(B2:B52)-I9" ' Formel för min-värde för X-axeln
            .Range("I3").formula = "=MAX(B2:B52)+I9" ' Formel för max-värde för X-axeln
            .Range("I4").formula = "=MIN(C2:C52)-I9" ' Formel för min-värde för Y-axeln
            .Range("I5").formula = "=MAX(C2:C52)+I9" ' Formel för max-värde för Y-axeln
            .Range("I6").formula = "=MEDIAN(B2:B52)" ' Formel för median-värde för X-axeln
            .Range("I7").formula = "=MEDIAN(C2:C52)" ' Formel för median-värde för Y-axeln
            .Range("I9").value = 0.02 ' Justering för korsningspunkt
        End With

        ' Läs axelvärden från kolumn I
        xMin = chartSheet.Range("I2").value
        xMax = chartSheet.Range("I3").value
        yMin = chartSheet.Range("I4").value
        yMax = chartSheet.Range("I5").value
        horizontalCrossing = chartSheet.Range("I6").value
        verticalCrossing = chartSheet.Range("I7").value

        ' Justera axlar i diagrammet
        With embeddedChart.chart
            ' Ställ in skalor och korsningspunkter för x- och y-axlar
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

        ' Immediate Response for Success
        Debug.Print "Axlar justerade med formler i Excel."
    Else
        MsgBox "Inget PowerPoint-diagram hittades på sliden.", vbExclamation
    End If
End Sub

