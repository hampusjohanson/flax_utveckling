Attribute VB_Name = "labels_left"
Sub ForceAlignLeftForFirstChart()
    Dim slide As slide
    Dim shp As shape
    Dim ch As chart

    ' Get the active slide
    Set slide = ActiveWindow.View.slide

    ' Find the first chart on the slide
    For Each shp In slide.Shapes
        If shp.hasChart Then
            Set ch = shp.chart
            shp.Select ' Select chart shape
            Exit For
        End If
    Next shp

    ' If no chart found, exit
    If ch Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Activate the chart for modifications
    ch.ChartArea.Select
    DoEvents

    ' Try executing a command to align text left
    On Error Resume Next
    CommandBars.ExecuteMso "AlignLeftText" ' Might need tweaking
    On Error GoTo 0
End Sub

