Attribute VB_Name = "Lines_3"
Sub Lines_3()

    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim lineWidth As Single ' Line width variable
    Dim chartIndex As Integer
    Dim SeriesIndex As Integer

    ' Prompt the user for the line width
    lineWidth = InputBox("Enter the desired line width for the series:", "Line Width", 1.75)
    
    ' Ensure the user entered a valid number greater than 0
    If lineWidth <= 0 Then
        MsgBox "Please enter a valid number greater than 0."
        Exit Sub
    End If

    ' === Loop through all charts on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    chartIndex = 1 ' Track chart number
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart

            ' Check if the chart has at least 12 series
            If chartObject.SeriesCollection.count >= 12 Then
                For SeriesIndex = 1 To 12
                    Set series = chartObject.SeriesCollection(SeriesIndex)

                    ' Apply the line weight
                    On Error Resume Next
                    With series.Format.line
                        .Weight = lineWidth ' Set the line weight
                    End With
                    If Err.Number <> 0 Then
                        Debug.Print "Error applying line weight to series " & SeriesIndex & " in chart " & chartIndex & ": " & Err.Description
                        Err.Clear
                    Else
                        Debug.Print "Line weight successfully applied to series " & SeriesIndex & " in chart " & chartIndex
                    End If
                    On Error GoTo 0
                Next SeriesIndex
            Else
                Debug.Print "The chart does not have 12 series."
                MsgBox "The chart does not have at least 12 series. Please check the chart.", vbExclamation
            End If
            chartIndex = chartIndex + 1 ' Move to the next chart
        End If
    Next chartShape

End Sub

