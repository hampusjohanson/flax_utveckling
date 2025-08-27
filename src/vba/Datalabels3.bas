Attribute VB_Name = "Datalabels3"
Sub DataLabels3()
    Dim pptApp As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pptChart As Object
    Dim ser As Object
    Dim i As Integer
    Dim lbl As Object
    Dim labelValue As String
    Dim excelSheet As Object
    Dim xMedian As Double
    Dim adjustmentFactor As Double

    ' H�mta PowerPoint och aktiv slide
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.slide

    ' Hitta diagrammet p� sliden (f�rsta diagrammet p� sliden)
    For Each pptShape In pptSlide.Shapes
        If pptShape.hasChart Then
            Set pptChart = pptShape.chart
            Exit For
        End If
    Next pptShape

    If pptChart Is Nothing Then
        MsgBox "Inget diagram hittades p� denna slide.", vbExclamation
        Exit Sub
    End If

    ' H�mta f�rsta serien i diagrammet
    Set ser = pptChart.SeriesCollection(1)

    ' H�mta den inb�ddade Excel-tabellen fr�n diagrammet
    Set excelSheet = pptChart.chartData.Workbook.Sheets(1) ' Det f�rsta bladet i den inb�ddade Excel-tabellen

    ' Ber�kna medianen f�r X-v�rden (och justeringsfaktor om det beh�vs)
    ' (Det h�r �r ett exempel p� hur medianen kan ber�knas, justera efter behov)
    xMedian = Median(ser) ' Du kan implementera en medianfunktion som passar din data
    adjustmentFactor = 10 ' Anpassa justeringsfaktorn som du �nskar

    ' Ta bort alla befintliga dataetiketter
    ser.HasDataLabels = False

    ' L�gg till nya dataetiketter
    ser.HasDataLabels = True

    ' Iterera genom alla datapunkter i serien och justera etiketterna
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' H�mta v�rdet fr�n kolumn A f�r dataetiketten
        labelValue = excelSheet.Cells(i + 1, 1).value ' Kolumn A, rad i (l�gg till 1 eftersom Excel �r 1-baserat)

        ' Tilldela etikettens text fr�n kolumn A
        lbl.text = labelValue

        ' Placera etiketten alltid ovanf�r datapunkten
        lbl.Position = xlLabelPositionAbove ' S�tt etiketten ovanf�r datapunkten

        ' �ndra textformat f�r etiketten
        With lbl
            .Font.Name = "Arial" ' Font: Arial
            .Font.size = 7 ' Fontstorlek: 7
            .Font.color = RGB(17, 21, 66) ' Textf�rg: RGB(17, 21, 66)
        End With
    Next i

End Sub

' Funktion f�r att ber�kna medianen (exempel)
Function Median(ser As Object) As Double
    Dim values() As Double
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double

    ' L�gg alla X-v�rden i en array
    ReDim values(ser.Points.count - 1)
    For i = 1 To ser.Points.count
        values(i - 1) = ser.Points(i).left ' Eller anv�nd annan data f�r X-v�rdena
    Next i

    ' Sortera v�rdena
    For i = 0 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) > values(j) Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i

    ' Ber�kna medianen
    If UBound(values) Mod 2 = 0 Then
        ' Om j�mnt antal element, ta medelv�rdet av de tv� mittersta
        Median = (values(UBound(values) \ 2) + values(UBound(values) \ 2 + 1)) / 2
    Else
        ' Om udda antal element, ta mittv�rdet
        Median = values(UBound(values) \ 2)
    End If
End Function

