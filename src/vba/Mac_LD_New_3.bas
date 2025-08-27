Attribute VB_Name = "Mac_LD_New_3"
Sub LD_3()
    ' Justera axlar
    
    Dim pptSlide As slide
    Dim embeddedChart As shape
    Dim chartWorkbook As Object ' Excel arbetsboken f�r diagramdata
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

    ' Hitta diagrammet p� sliden
    Set pptSlide = ActiveWindow.View.slide

    ' S�k igenom alla objekt p� sliden f�r att hitta ett diagram
    For shapeIndex = 1 To pptSlide.Shapes.count
        shapeType = pptSlide.Shapes(shapeIndex).Type
        If shapeType = msoChart Then
            Set embeddedChart = pptSlide.Shapes(shapeIndex)
            Exit For ' Hitta och avsluta loopen
        End If
    Next shapeIndex

    ' Om ett diagram finns, forts�tt
    If Not embeddedChart Is Nothing Then
        ' H�mta Excel arbetsboken och arket f�r diagrammet
        Set chartWorkbook = embeddedChart.chart.chartData.Workbook
        Set chartSheet = chartWorkbook.Sheets(1)

        ' Skapa formler i kolumn I f�r att ber�kna min- och max-v�rden samt medianer
        With chartSheet
            .Range("I2").formula = "=MIN(B2:B52)-I9" ' Formel f�r min-v�rde f�r X-axeln
            .Range("I3").formula = "=MAX(B2:B52)+I9" ' Formel f�r max-v�rde f�r X-axeln
            .Range("I4").formula = "=MIN(C2:C52)-I9" ' Formel f�r min-v�rde f�r Y-axeln
            .Range("I5").formula = "=MAX(C2:C52)+I9" ' Formel f�r max-v�rde f�r Y-axeln
            .Range("I6").formula = "=MEDIAN(B2:B52)" ' Formel f�r median-v�rde f�r X-axeln
            .Range("I7").formula = "=MEDIAN(C2:C52)" ' Formel f�r median-v�rde f�r Y-axeln
            .Range("I9").value = 0.02 ' Justering f�r korsningspunkt
        End With

        ' L�s axelv�rden fr�n kolumn I
        xMin = chartSheet.Range("I2").value
        xMax = chartSheet.Range("I3").value
        yMin = chartSheet.Range("I4").value
        yMax = chartSheet.Range("I5").value
        horizontalCrossing = chartSheet.Range("I6").value
        verticalCrossing = chartSheet.Range("I7").value

        ' Justera axlar i diagrammet
        With embeddedChart.chart
            ' St�ll in skalor och korsningspunkter f�r x- och y-axlar
            .Axes(xlCategory).MinimumScale = xMin
            .Axes(xlCategory).MaximumScale = xMax
            .Axes(xlCategory).CrossesAt = horizontalCrossing
            .Axes(xlValue).MinimumScale = yMin
            .Axes(xlValue).MaximumScale = yMax
            .Axes(xlValue).CrossesAt = verticalCrossing

            ' Anpassa utseendet p� axellinjerna
            With .Axes(xlCategory).Format.line
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66) ' F�rg
                .DashStyle = msoLineLongDash ' L�ng streckad linje
                .Weight = 0.25 ' Linjetjocklek
            End With

            With .Axes(xlValue).Format.line
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66) ' F�rg
                .DashStyle = msoLineLongDash ' L�ng streckad linje
                .Weight = 0.25 ' Linjetjocklek
            End With
        End With

        ' Immediate Response for Success
        Debug.Print "Axlar justerade med formler i Excel."
    Else
        MsgBox "Inget PowerPoint-diagram hittades p� sliden.", vbExclamation
    End If
End Sub

