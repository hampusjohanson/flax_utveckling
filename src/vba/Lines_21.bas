Attribute VB_Name = "Lines_21"
Sub Lines_21()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim series As series
    Dim visibleSeriesColors() As Long
    Dim visibleSeriesCount As Integer
    Dim i As Integer
    Dim tableShapes As Object
    Dim tbl As table
    Dim rowColMap As Collection
    Dim key As Variant

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loop through all shapes to find the first chart
    For Each shape In pptSlide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart ' Set the first chart found
            Exit For ' Exit loop after finding the first chart
        End If
    Next shape

    ' If no chart found, show an error and exit
    If chart Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Initialize variables
    ReDim visibleSeriesColors(1 To chart.SeriesCollection.count)
    visibleSeriesCount = 0

    ' Collect the colors of all visible series
    For i = 1 To chart.SeriesCollection.count
        Set series = chart.SeriesCollection(i)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleSeriesCount = visibleSeriesCount + 1
            visibleSeriesColors(visibleSeriesCount) = series.Format.line.ForeColor.RGB ' Store the color
        End If
    Next i

    ' Exit if no visible series were found
    If visibleSeriesCount = 0 Then
        MsgBox "No visible series found in the chart.", vbExclamation
        Exit Sub
    End If

    ' Initialize rowColMap as a Collection
    Set rowColMap = New Collection

    ' Map the table cell locations to the corresponding series colors
    rowColMap.Add Array("Brand_List_1", 1, 1, 1) ' Table 1, Row 1, Column 1 -> Series 1
    rowColMap.Add Array("Brand_List_1", 2, 1, 2) ' Table 1, Row 2, Column 1 -> Series 2
    rowColMap.Add Array("Brand_List_1", 3, 1, 3) ' Table 1, Row 3, Column 1 -> Series 3
    rowColMap.Add Array("Brand_List_2", 1, 1, 4) ' Table 2, Row 1, Column 1 -> Series 4
    rowColMap.Add Array("Brand_List_2", 2, 1, 5) ' Table 2, Row 2, Column 1 -> Series 5
    rowColMap.Add Array("Brand_List_2", 3, 1, 6) ' Table 2, Row 3, Column 1 -> Series 6
    rowColMap.Add Array("Brand_List_3", 1, 1, 7) ' Table 3, Row 1, Column 1 -> Series 7
    rowColMap.Add Array("Brand_List_3", 2, 1, 8) ' Table 2, Row 2, Column 1 -> Series 8

    ' Loop through the table map and update the fill colors
    For Each key In rowColMap
        Dim parts() As Variant
        parts = key
        
        ' Skip if the series index exceeds the visible count
        If parts(3) > visibleSeriesCount Then Exit For

        ' Find the table by name
        On Error Resume Next
        Set shape = pptSlide.Shapes(parts(0))
        On Error GoTo 0

        If Not shape Is Nothing And shape.Type = msoTable Then
            ' Update the fill color of the specified cell
            Set tbl = shape.table
            tbl.cell(parts(1), parts(2)).shape.Fill.ForeColor.RGB = visibleSeriesColors(parts(3))
        End If
    Next key
End Sub

