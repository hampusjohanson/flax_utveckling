Attribute VB_Name = "Substance_border_1"
Sub Substance_border_1()
    Dim slide As slide
    Dim chart As chart
    Dim series As series

    ' Ensure the active slide contains at least one shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a slide that contains a chart.", vbExclamation
        Exit Sub
    End If

    Set slide = ActiveWindow.View.slide
    
    ' Loop through all shapes to find the first chart
    Dim chartFound As Boolean
    chartFound = False

    For Each shape In slide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart
            chartFound = True
            shape.Select ' Select the chart
            Exit For
        End If
    Next shape

    ' If no chart was found
    If Not chartFound Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Ensure there is exactly one series in the chart
    If chart.SeriesCollection.count <> 1 Then
        MsgBox "The chart must have exactly one series.", vbExclamation
        Exit Sub
    End If

    ' Get the first (and only) series
    Set series = chart.SeriesCollection(1)

    ' Remove the outline (border) for the entire series (points)
    With series.Format.line
        .visible = msoFalse ' No border/outline
    End With

    MsgBox "Outline (border) removed from the series data points.", vbInformation
End Sub


