Attribute VB_Name = "DataLabels6"
Sub DataLabels6()
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
    Dim xMin As Double
    Dim xMax As Double
    Dim leftThreshold As Double
    Dim rightThreshold As Double

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

    ' Ber�kna medianen f�r X-v�rdena
    xMedian = Median(ser)
    xMin = Min(ser)
    xMax = Max(ser)

    ' Definiera tr�skelv�rden f�r v�nster, centralt och h�ger
    leftThreshold = xMedian - ((xMax - xMin) / 4)
    rightThreshold = xMedian + ((xMax - xMin) / 4)

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

        ' Justera etikettens position baserat p� X-v�rdet
        If ser.Points(i).left < leftThreshold Then
            lbl.Position = xlLabelPositionLeft ' S�tt etiketten till v�nster om datapunkten
        ElseIf ser.Points(i).left > rightThreshold Then
            lbl.Position = xlLabelPositionRight ' S�tt etiketten till h�ger om datapunkten
        Else
            lbl.Position = xlLabelPositionAbove ' S�tt etiketten ovanf�r datapunkten (centrerad)
        End If

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

' Funktion f�r att hitta det minsta X-v�rdet (v�nster)
Function Min(ser As Object) As Double
    Dim minValue As Double
    minValue = ser.Points(1).left
    For i = 2 To ser.Points.count
        If ser.Points(i).left < minValue Then
            minValue = ser.Points(i).left
        End If
    Next i
    Min = minValue
End Function

' Funktion f�r att hitta det st�rsta X-v�rdet (h�ger)
Function Max(ser As Object) As Double
    Dim maxValue As Double
    maxValue = ser.Points(1).left
    For i = 2 To ser.Points.count
        If ser.Points(i).left > maxValue Then
            maxValue = ser.Points(i).left
        End If
    Next i
    Max = maxValue
End Function



