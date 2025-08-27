Attribute VB_Name = "Lines_25"
Sub Lines_25()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim tbl As table
    Dim series As series
    Dim visibleCount As Integer
    Dim i As Integer

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Debugging: Log active slide
    Debug.Print "Active Slide: " & pptSlide.SlideIndex

    ' Loop through all shapes to find the first chart
    For Each shape In pptSlide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart
            Debug.Print "Chart found on slide."
            Exit For
        End If
    Next shape

    ' If no chart found, show an error and exit
    If chart Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Debug.Print "No chart found on the slide."
        Exit Sub
    End If

    ' Initialize visible brand count (excluding the last series)
    visibleCount = 0

    ' Count visible brands
    For i = 1 To chart.SeriesCollection.count - 1 ' Exclude the last series
        Set series = chart.SeriesCollection(i)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleCount = visibleCount + 1
        End If
    Next i

    ' Debugging: Log total visible brand count
    Debug.Print "Total visible brands: " & visibleCount

    ' Find the table named "Brand_List_3"
    On Error Resume Next
    Set shape = pptSlide.Shapes("Brand_List_3")
    On Error GoTo 0

    ' If the table is found, apply logic for visibleCount
    If Not shape Is Nothing And shape.Type = msoTable Then
        Set tbl = shape.table
        Debug.Print "Processing table: Brand_List_3"

        If visibleCount = 8 Then
            ' Set column 1, row 3 to no color
            tbl.cell(3, 1).shape.Fill.visible = msoFalse
            Debug.Print "Set column 1, row 3 to no color for visibleCount = 8"
        End If

    Else
        Debug.Print "Table 'Brand_List_3' not found. Skipping."
    End If
End Sub

