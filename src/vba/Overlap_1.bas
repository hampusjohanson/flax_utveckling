Attribute VB_Name = "Overlap_1"
Sub A_Overlap_Onetime()
    Dim pptApp As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pptChart As Object
    Dim ser As Object
    Dim OverlappingLabels As Collection
    Dim allLabels As Collection
    Dim lbl As Object
    Dim i As Integer
    Dim labelKey As String
    Dim nonOverlapping As Collection
    Dim result As String

    ' H�mta PowerPoint och aktiv slide
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.slide

    ' Hitta diagrammet p� sliden
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

    ' H�mta f�rsta serien
    Set ser = pptChart.SeriesCollection(1)

    ' Samla alla etiketter och exkludera ogiltiga
    Set allLabels = New Collection
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Kontrollera om etiketten ska exkluderas
        If IsValidLabel(lbl) Then
            allLabels.Add lbl, "P" & i
        End If
    Next i

    ' Hitta etiketter som �verlappar
    Set OverlappingLabels = FindOverlappingLabelsFiltered(ser, allLabels)

    ' Identifiera etiketter som aldrig �verlappar
    Set nonOverlapping = New Collection
    For i = 1 To allLabels.count
        labelKey = "P" & i
        On Error Resume Next
        ' Om etiketten inte finns i overlappingLabels, l�gg till i nonOverlapping
        If OverlappingLabels(labelKey) Is Nothing Then
            nonOverlapping.Add allLabels(labelKey), labelKey
        End If
        On Error GoTo 0
    Next i

    ' Skapa rapport
    result = "Rapport f�r etiketter (filtret till�mpat):" & vbNewLine & vbNewLine
    result = result & "Antal �verlappande etiketter: " & OverlappingLabels.count & vbNewLine
    result = result & "Antal etiketter utan �verlapp: " & nonOverlapping.count & vbNewLine & vbNewLine

    result = result & "�verlappande etiketter:" & vbNewLine
    For i = 1 To OverlappingLabels.count
        Set lbl = OverlappingLabels(i)
        result = result & lbl.text & vbNewLine
    Next i

    result = result & vbNewLine & "Etiketter utan �verlapp:" & vbNewLine
    For i = 1 To nonOverlapping.count
        Set lbl = nonOverlapping(i)
        result = result & lbl.text & vbNewLine
    Next i

    ' Visa rapport
    MsgBox result, vbInformation
End Sub

Function IsValidLabel(lbl As Object) As Boolean
    ' Kontrollera om etiketten �r giltig
    If lbl.text = "False" Or lbl.text = "Falskt" Or lbl.text = "" Then
        IsValidLabel = False
    Else
        IsValidLabel = True
    End If
End Function

Function FindOverlappingLabelsFiltered(ser As Object, validLabels As Collection) As Collection
    Dim i As Integer, j As Integer
    Dim lbl1 As Object, lbl2 As Object
    Dim toleranceX As Double, toleranceY As Double
    Dim OverlappingLabels As New Collection
    Dim labelKey As String

    ' Tolerans f�r att definiera �verlapp
    toleranceX = 30 ' Horisontellt
    toleranceY = 15 ' Vertikalt

    ' Iterera genom alla etikettpar fr�n giltiga etiketter
    For i = 1 To validLabels.count
        Set lbl1 = validLabels(i)
        For j = i + 1 To validLabels.count
            Set lbl2 = validLabels(j)

            ' Kontrollera �verlapp
            If Abs(lbl1.left - lbl2.left) < toleranceX And Abs(lbl1.Top - lbl2.Top) < toleranceY Then
                ' L�gg till etikett 1 om den inte redan �r med
                labelKey = "P" & i
                On Error Resume Next
                OverlappingLabels.Add lbl1, labelKey
                On Error GoTo 0

                ' L�gg till etikett 2 om den inte redan �r med
                labelKey = "P" & jz
                On Error Resume Next
                OverlappingLabels.Add lbl2, labelKey
                On Error GoTo 0
            End If
        Next j
    Next i

    ' Returnera unika etiketter som �verlappar
    Set FindOverlappingLabelsFiltered = OverlappingLabels
End Function

