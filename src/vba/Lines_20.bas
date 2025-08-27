Attribute VB_Name = "Lines_20"
Sub Lines_20()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim series As series
    Dim brandList1 As shape
    Dim visibleSeriesColor As Long
    Dim i As Integer
    Dim foundVisibleSeries As Boolean

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
    foundVisibleSeries = False

    ' Loop through all series to find the first visible series
    For i = 1 To chart.SeriesCollection.count
        Set series = chart.SeriesCollection(i)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleSeriesColor = series.Format.line.ForeColor.RGB ' Get the line color of the first visible series
            foundVisibleSeries = True
            Exit For
        End If
    Next i

    ' If no visible series found, exit
    If Not foundVisibleSeries Then
        MsgBox "No visible series found in the chart.", vbExclamation
        Exit Sub
    End If

    ' Find the table named "Brand_List_1"
    On Error Resume Next
    Set brandList1 = pptSlide.Shapes("Brand_List_1")
    On Error GoTo 0

    If brandList1 Is Nothing Or brandList1.Type <> msoTable Then
        MsgBox "'Brand_List_1' table not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Set the fill color of column 1, row 1 in "Brand_List_1" to match the series line color
    brandList1.table.cell(1, 1).shape.Fill.ForeColor.RGB = visibleSeriesColor

    ' Notify the user
    MsgBox "The fill color of Brand_List_1 column 1 row 1 has been updated to match the first visible series line color."
End Sub

