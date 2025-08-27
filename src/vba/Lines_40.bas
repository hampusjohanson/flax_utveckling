Attribute VB_Name = "Lines_40"
Sub Lines_40()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim SeriesIndex As Integer
    Dim chartCount As Integer
    Dim hiddenSeriesCount As Integer

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    chartCount = 0
    hiddenSeriesCount = 0

    ' Loop through all shapes on the slide to find charts
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            chartCount = chartCount + 1 ' Count charts

            ' Loop through all series in the chart
            For SeriesIndex = 1 To chartObject.SeriesCollection.count
                Set series = chartObject.SeriesCollection(SeriesIndex)

                ' Check if the series name is "FALSKT" or "FALSE" (case-insensitive)
                If LCase(series.Name) = "falskt" Or LCase(series.Name) = "false" Then
                    On Error Resume Next
                    series.Format.line.visible = msoFalse  ' Hide the line
                    series.MarkerStyle = xlMarkerStyleNone ' Hide markers
                    Debug.Print "Series hidden in Chart " & chartCount & ": " & series.Name
                    hiddenSeriesCount = hiddenSeriesCount + 1
                    On Error GoTo 0
                End If
            Next SeriesIndex
        End If
    Next chartShape

    ' Show result message
    If chartCount = 0 Then
        MsgBox "No charts found on this slide.", vbExclamation
    ElseIf hiddenSeriesCount = 0 Then
       
    Else
       
    End If
End Sub

