Attribute VB_Name = "AA_Series_set_1_as_2"
Sub AA_Series_set_1_as_2()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim series1 As series

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

    ' Check if the chart was found
    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Access Series1 in the chart ===
    On Error Resume Next
    Set series1 = chartShape.chart.SeriesCollection(1)
    On Error GoTo 0

    If series1 Is Nothing Then
        MsgBox "Series1 not found in the chart.", vbExclamation
        Exit Sub
    End If

    ' === Set the data labels' font color ===
    With series1
        .ApplyDataLabels
        .DataLabels.Font.color = RGB(255, 255, 255) ' Font color: White
    End With

    ' === Set the fill color of the series ===
    With series1.Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(112, 113, 140) ' Fill color: Hex #70718C
        .Solid
    End With

   
End Sub

