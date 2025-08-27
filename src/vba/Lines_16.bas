Attribute VB_Name = "Lines_16"
Sub Lines_16()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim totalSeries As Integer
    Dim visibleBrandCount As Integer
    Dim currentVisibleIndex As Integer
    Dim i As Integer
    Dim tbl As table
    Dim series As series

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

    ' Get total number of series in the chart
    totalSeries = chart.SeriesCollection.count

    ' Initialize visible series count
    visibleBrandCount = 0
    currentVisibleIndex = 0

    ' Find the table named "Brand_List_3"
    On Error Resume Next
    Set shape = pptSlide.Shapes("Brand_List_3")
    On Error GoTo 0

    If shape Is Nothing Or shape.Type <> msoTable Then
        MsgBox "'Brand_List_3' table not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Set the table reference
    Set tbl = shape.table

    ' Loop through all series and count visible ones
    For i = 1 To totalSeries
        Set series = chart.SeriesCollection(i)

        ' Check if the series is visible (both line and markers)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleBrandCount = visibleBrandCount + 1

            ' Dynamically handle the 7th and 8th visible brands
            If visibleBrandCount = 7 Then
                tbl.cell(1, 2).shape.TextFrame.textRange.text = series.Name ' Set the 7th visible brand name
            ElseIf visibleBrandCount = 8 Then
                tbl.cell(2, 2).shape.TextFrame.textRange.text = series.Name ' Set the 8th visible brand name
            End If
        End If
    Next i

    ' Handle clearing text based on visible brands
    If visibleBrandCount = 7 Then
        tbl.cell(2, 2).shape.TextFrame.textRange.text = "" ' Clear row 2, column 2
        tbl.cell(3, 2).shape.TextFrame.textRange.text = "" ' Clear row 3, column 2
    ElseIf visibleBrandCount = 8 Then
        tbl.cell(3, 2).shape.TextFrame.textRange.text = "" ' Clear row 3, column 2
    End If

  
End Sub

