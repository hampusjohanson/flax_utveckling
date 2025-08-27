Attribute VB_Name = "AA_Series_set_2_as_3"
Sub AA_Series_set_2_as_3()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim series2 As series

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

    ' === Access Series2 in the chart ===
    On Error Resume Next
    Set series2 = chartShape.chart.SeriesCollection(2)
    On Error GoTo 0

    If series2 Is Nothing Then
        MsgBox "Series2 not found in the chart.", vbExclamation
        Exit Sub
    End If

    ' === Set the data labels' font color ===
    With series2
        .ApplyDataLabels
        .DataLabels.Font.color = RGB(255, 255, 255) ' Font color: White
    End With

    ' === Set the fill color of the series ===
    With series2.Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(158, 159, 177) ' Fill color: RGB(158, 159, 177) or Hex #9E9FB1
        .Solid
    End With

    
End Sub

