Attribute VB_Name = "Capital_4"
Sub Mac_Cap_Labels1()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim rowIndex As Integer
    Dim dataLabel As Object

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Aktuell slide h�mtad"

    ' Hitta diagrammet p� sliden
    On Error GoTo ErrorHandler
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades p� sliden.", vbCritical
        Debug.Print "Inget diagram hittades"
        Exit Sub
    End If
    
    Debug.Print "Diagram hittades p� sliden"

    ' F�rs�k att �ppna diagrammets datak�lla
    On Error Resume Next
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo ErrorHandler
    
    If chartDataWorkbook Is Nothing Then
        MsgBox "Diagrammet har ingen datak�lla.", vbCritical
        Debug.Print "Diagrammet har ingen datak�lla"
        Exit Sub
    End If

    Debug.Print "Datak�llan �ppnad"
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
                dataLabel.Position = xlLabelPositionRight ' V�nsterst�llda etiketter
                Debug.Print "Etikett tillagd p� punkt " & rowIndex
            End If
        Next rowIndex
    End With

    ' St�ng diagrammets datak�lla
    chartShape.chart.chartData.Workbook.Close
    Debug.Print "Datak�lla st�ngd"

    Exit Sub

ErrorHandler:
    MsgBox "Ett fel intr�ffade: " & Err.Description, vbCritical
    Debug.Print "Fel: " & Err.Description
End Sub

