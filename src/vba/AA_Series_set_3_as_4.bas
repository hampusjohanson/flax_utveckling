Attribute VB_Name = "AA_Series_set_3_as_4"
Sub AA_Series_set_3_as_4()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim series3 As series

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

    ' === Access Series3 in the chart ===
    On Error Resume Next
    Set series3 = chartShape.chart.SeriesCollection(3)
    On Error GoTo 0

    If series3 Is Nothing Then
        MsgBox "Series3 not found in the chart.", vbExclamation
        Exit Sub
    End If

    ' === Set the data labels' font color ===
    With series3
        .ApplyDataLabels
        .DataLabels.Font.color = RGB(17, 21, 66) ' Font color: RGB(17, 21, 66)
    End With

    ' === Set the fill color of the series ===
    With series3.Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(231, 232, 237) ' Fill color: Hex #E7E8ED
        .Solid
    End With

    
End Sub

