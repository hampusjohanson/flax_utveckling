Attribute VB_Name = "AA_Count_boxes_series_A5"
Sub AA_Count_boxes_series_A5()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim series1 As series
    Dim series2 As series
    Dim val1 As Double
    Dim val2 As Double
    Dim sumValue As Double
    Dim percentageText As String
    Dim leftieBox As shape

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Locate the first chart on the slide ===
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape ' Use the first chart found
            Exit For
        End If
    Next shape

    ' === Get first two series ===
    On Error Resume Next
    Set series1 = chartShape.chart.SeriesCollection(1)
    Set series2 = chartShape.chart.SeriesCollection(2)
    On Error GoTo 0

    ' === Ensure there are at least 5 data points in both series ===
    If series1 Is Nothing Or series2 Is Nothing Then Exit Sub
    If series1.Points.count < 5 Or series2.Points.count < 5 Then Exit Sub

    ' === Extract the fifth value from each series ===
    val1 = series1.values(5)
    val2 = series2.values(5)

    ' === Handle PowerPoint formatting (Convert if needed) ===
    If val1 > 1 Then val1 = val1 / 100 ' Convert from 21 to 0.21 if necessary
    If val2 > 1 Then val2 = val2 / 100 ' Convert from 73 to 0.73 if necessary

    ' === Sum the values ===
    sumValue = val1 + val2

    ' === Convert to percentage format ===
    percentageText = Format(sumValue, "0%")

    ' === Find "Leftie_5" box ===
    Set leftieBox = Nothing
    For Each shape In pptSlide.Shapes
        If shape.Name = "Leftie_5" Then
            Set leftieBox = shape
            Exit For
        End If
    Next shape

    ' === Exit if Leftie_5 was not found ===
    If leftieBox Is Nothing Then Exit Sub

    ' === Paste the percentage into Leftie_5 ===
    leftieBox.TextFrame.textRange.text = percentageText

    ' === Debugging Output ===
    Debug.Print "First value in Series 1: " & val1 * 100 & "%"
    Debug.Print "First value in Series 2: " & val2 * 100 & "%"
    Debug.Print "Sum (as percentage): " & percentageText
    Debug.Print "Percentage pasted into Leftie_5."

End Sub

