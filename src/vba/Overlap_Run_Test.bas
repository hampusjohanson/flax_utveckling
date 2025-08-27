Attribute VB_Name = "Overlap_Run_Test"
Public overlapCount As Integer ' Global variabel för att hålla antal överlapp

Sub RunDataLabels1()
    ' Kör DataLabels1 makrot via Application.Run om det finns i en annan modul
    Application.Run "DataLabels1"
End Sub

Sub RunDataLabelsAndCountOverlaps()
    ' Variabel för att hålla sammanfattningen av resultaten
    Dim resultSummary As String

    ' Initialisera resultSummary
    resultSummary = "Överlapp per makro:" & vbNewLine & vbNewLine

    ' Kör DataLabels1 och räkna överlappen
    Application.Run "DataLabels1" ' Kör DataLabels1 makrot
    resultSummary = resultSummary & "DataLabels1: " & CountOverlappingLabels() & " överlapp." & vbNewLine

    ' Kör DataLabels2 och räkna överlappen
    Application.Run "DataLabels2" ' Kör DataLabels2 makrot
    resultSummary = resultSummary & "DataLabels2: " & CountOverlappingLabels() & " överlapp." & vbNewLine

    ' Kör DataLabels3 och räkna överlappen
    Application.Run "DataLabels3" ' Kör DataLabels3 makrot
    resultSummary = resultSummary & "DataLabels3: " & CountOverlappingLabels() & " överlapp." & vbNewLine
    
        ' Kör DataLabels4 och räkna överlappen
    Application.Run "DataLabels4" ' Kör DataLabels4 makrot
    resultSummary = resultSummary & "DataLabels4: " & CountOverlappingLabels() & " överlapp." & vbNewLine

    ' Kör DataLabels5 och räkna överlappen
    Application.Run "DataLabels5" ' Kör DataLabels5 makrot
    resultSummary = resultSummary & "DataLabels5: " & CountOverlappingLabels() & " överlapp." & vbNewLine

    ' Kör DataLabels6 och räkna överlappen
    Application.Run "DataLabels6" ' Kör DataLabels6 makrot
    resultSummary = resultSummary & "DataLabels6: " & CountOverlappingLabels() & " överlapp." & vbNewLine


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

    ' Hämta PowerPoint och aktiv slide
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.slide

    ' Hitta diagrammet på sliden
    For Each pptShape In pptSlide.Shapes
        If pptShape.hasChart Then
            Set pptChart = pptShape.chart
            Exit For
        End If
    Next pptShape

    If pptChart Is Nothing Then
        MsgBox "Inget diagram hittades på denna slide.", vbExclamation
        Exit Function
    End If

    ' Hämta första serien
    Set ser = pptChart.SeriesCollection(1)

    ' Samla alla etiketter och exkludera ogiltiga
    Set allLabels = New Collection
    For i = 1 To ser.Points.count
        Set lbl = ser.Points(i).dataLabel

        ' Kontrollera etikettens textvärde innan vi lägger till den i allLabels
        If IsValidLabel(lbl.text) Then
            allLabels.Add lbl, "P" & i
        End If
    Next i

    ' Hitta etiketter som överlappar
    Set OverlappingLabels = FindOverlappingLabelsFiltered(ser, allLabels)

    ' Räkna antal överlappande etiketter
    overlapCount = OverlappingLabels.count ' Uppdatera global variabel

    ' Återvänd antal överlappande etiketter
    CountOverlappingLabels = overlapCount
End Function

' Funktion för att kontrollera om etiketten är giltig
Function IsValidLabel(value As String) As Boolean
    ' Kontrollera om etiketten är ogiltig
    If Trim(value) = "" Or value = "False" Or value = "Falskt" Then
        IsValidLabel = False ' Ogiltig
    Else
        IsValidLabel = True ' Giltig
    End If
End Function

