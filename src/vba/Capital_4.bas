Attribute VB_Name = "Capital_4"
Sub Mac_Cap_Labels1()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim rowIndex As Integer
    Dim dataLabel As Object

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Aktuell slide hämtad"

    ' Hitta diagrammet på sliden
    On Error GoTo ErrorHandler
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbCritical
        Debug.Print "Inget diagram hittades"
        Exit Sub
    End If
    
    Debug.Print "Diagram hittades på sliden"

    ' Försök att öppna diagrammets datakälla
    On Error Resume Next
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo ErrorHandler
    
    If chartDataWorkbook Is Nothing Then
        MsgBox "Diagrammet har ingen datakälla.", vbCritical
        Debug.Print "Diagrammet har ingen datakälla"
        Exit Sub
    End If

    Debug.Print "Datakällan öppnad"
    Set chartSheet = chartDataWorkbook.Sheets(1)

    ' Aktivera datalabels
    With chartShape.chart.SeriesCollection(1)
        .ApplyDataLabels xlDataLabelsShowValue
        Debug.Print "DataLabels applicerade"
        For rowIndex = 2 To 13 ' A2 till A13
            If rowIndex - 1 <= .Points.count Then
                ' Se till att vi har en punkt att tilldela etikett till
                Set dataLabel = .Points(rowIndex - 1).dataLabel
                dataLabel.text = chartSheet.Cells(rowIndex, 1).value ' Kolumn A
                dataLabel.Font.Name = "Arial"
                dataLabel.Font.size = 10
                dataLabel.Font.color = RGB(17, 21, 66)
                dataLabel.Position = xlLabelPositionRight ' Vänsterställda etiketter
                Debug.Print "Etikett tillagd på punkt " & rowIndex
            End If
        Next rowIndex
    End With

    ' Stäng diagrammets datakälla
    chartShape.chart.chartData.Workbook.Close
    Debug.Print "Datakälla stängd"

    Exit Sub

ErrorHandler:
    MsgBox "Ett fel inträffade: " & Err.Description, vbCritical
    Debug.Print "Fel: " & Err.Description
End Sub

