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

    ' Hämta PowerPoint och aktiv slide
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.slide

    ' Hitta diagrammet på sliden (första diagrammet på sliden)
    For Each pptShape In pptSlide.Shapes
        If pptShape.hasChart Then
            Set pptChart = pptShape.chart
            Exit For
        End If
    Next pptShape

    If pptChart Is Nothing Then
        MsgBox "Inget diagram hittades på denna slide.", vbExclamation
        Exit Sub
    End If

    ' Hämta första serien i diagrammet
    Set ser = pptChart.SeriesCollection(1)

    ' Hämta den inbäddade Excel-tabellen från diagrammet
    Set excelSheet = pptChart.chartData.Workbook.Sheets(1) ' Det första bladet i den inbäddade Excel-tabellen

    ' Beräkna medianen för X-värden (och justeringsfaktor om det behövs)
    ' (Det här är ett exempel på hur medianen kan beräknas, justera efter behov)
    xMedian = Median(ser) ' Du kan implementera en medianfunktion som passar din data
    adjustmentFactor = 10 ' Anpassa justeringsfaktorn som du önskar

    ' Ta bort alla befintliga dataetiketter
    ser.HasDataLabels = False

    ' Lägg till nya dataetiketter
    ser.HasDataLabels = True

    ' Iterera genom alla datapunkter i serien och justera etiketterna
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Hämta värdet från kolumn A för dataetiketten
        labelValue = excelSheet.Cells(i + 1, 1).value ' Kolumn A, rad i (lägg till 1 eftersom Excel är 1-baserat)

        ' Tilldela etikettens text från kolumn A
        lbl.text = labelValue

        ' Placera etiketten alltid ovanför datapunkten
        lbl.Position = xlLabelPositionAbove ' Sätt etiketten ovanför datapunkten

        ' Ändra textformat för etiketten
        With lbl
            .Font.Name = "Arial" ' Font: Arial
            .Font.size = 7 ' Fontstorlek: 7
            .Font.color = RGB(17, 21, 66) ' Textfärg: RGB(17, 21, 66)
        End With
    Next i

End Sub

' Funktion för att beräkna medianen (exempel)
Function Median(ser As Object) As Double
    Dim values() As Double
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double

    ' Lägg alla X-värden i en array
    ReDim values(ser.Points.count - 1)
    For i = 1 To ser.Points.count
        values(i - 1) = ser.Points(i).left ' Eller använd annan data för X-värdena
    Next i

    ' Sortera värdena
    For i = 0 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(i) > values(j) Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i

    ' Beräkna medianen
    If UBound(values) Mod 2 = 0 Then
        ' Om jämnt antal element, ta medelvärdet av de två mittersta
        Median = (values(UBound(values) \ 2) + values(UBound(values) \ 2 + 1)) / 2
    Else
        ' Om udda antal element, ta mittvärdet
        Median = values(UBound(values) \ 2)
    End If
End Function

