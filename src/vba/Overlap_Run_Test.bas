Attribute VB_Name = "Overlap_Run_Test"
Public overlapCount As Integer ' Global variabel f�r att h�lla antal �verlapp

Sub RunDataLabels1()
    ' K�r DataLabels1 makrot via Application.Run om det finns i en annan modul
    Application.Run "DataLabels1"
End Sub

Sub RunDataLabelsAndCountOverlaps()
    ' Variabel f�r att h�lla sammanfattningen av resultaten
    Dim resultSummary As String

    ' Initialisera resultSummary
    resultSummary = "�verlapp per makro:" & vbNewLine & vbNewLine

    ' K�r DataLabels1 och r�kna �verlappen
    Application.Run "DataLabels1" ' K�r DataLabels1 makrot
    resultSummary = resultSummary & "DataLabels1: " & CountOverlappingLabels() & " �verlapp." & vbNewLine

    ' K�r DataLabels2 och r�kna �verlappen
    Application.Run "DataLabels2" ' K�r DataLabels2 makrot
    resultSummary = resultSummary & "DataLabels2: " & CountOverlappingLabels() & " �verlapp." & vbNewLine

    ' K�r DataLabels3 och r�kna �verlappen
    Application.Run "DataLabels3" ' K�r DataLabels3 makrot
    resultSummary = resultSummary & "DataLabels3: " & CountOverlappingLabels() & " �verlapp." & vbNewLine
    
        ' K�r DataLabels4 och r�kna �verlappen
    Application.Run "DataLabels4" ' K�r DataLabels4 makrot
    resultSummary = resultSummary & "DataLabels4: " & CountOverlappingLabels() & " �verlapp." & vbNewLine

    ' K�r DataLabels5 och r�kna �verlappen
    Application.Run "DataLabels5" ' K�r DataLabels5 makrot
    resultSummary = resultSummary & "DataLabels5: " & CountOverlappingLabels() & " �verlapp." & vbNewLine

    ' K�r DataLabels6 och r�kna �verlappen
    Application.Run "DataLabels6" ' K�r DataLabels6 makrot
    resultSummary = resultSummary & "DataLabels6: " & CountOverlappingLabels() & " �verlapp." & vbNewLine


    ' Visa sammanfattningen av resultaten
    MsgBox resultSummary
End Sub

Function CountOverlappingLabels() As Integer
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
    Dim overlapCount As Integer

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
        Exit Function
    End If

    ' H�mta f�rsta serien
    Set ser = pptChart.SeriesCollection(1)

    ' Samla alla etiketter och exkludera ogiltiga
    Set allLabels = New Collection
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Kontrollera etikettens textv�rde innan vi l�gger till den i allLabels
        If IsValidLabel(lbl.text) Then
            allLabels.Add lbl, "P" & i
        End If
    Next i

    ' Hitta etiketter som �verlappar
    Set OverlappingLabels = FindOverlappingLabelsFiltered(ser, allLabels)

    ' R�kna antal �verlappande etiketter
    overlapCount = OverlappingLabels.count ' Uppdatera global variabel

    ' �terv�nd antal �verlappande etiketter
    CountOverlappingLabels = overlapCount
End Function

' Funktion f�r att kontrollera om etiketten �r giltig
Function IsValidLabel(value As String) As Boolean
    ' Kontrollera om etiketten �r ogiltig
    If Trim(value) = "" Or value = "False" Or value = "Falskt" Then
        IsValidLabel = False ' Ogiltig
    Else
        IsValidLabel = True ' Giltig
    End If
End Function

