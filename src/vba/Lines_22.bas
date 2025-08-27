Attribute VB_Name = "Lines_22"
Sub Lines_22()
      Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim series As series
    Dim visibleCount As Integer

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

    ' Process each table independently
    Call ProcessBrandTable(visibleCount, "Brand_List_1")
    Call ProcessBrandTable(visibleCount, "Brand_List_2")
    Call ProcessBrandTable(visibleCount, "Brand_List_3")
End Sub


Sub ProcessBrandTable(visibleCount As Integer, tableName As String)
    Dim shape As shape
    Dim tbl As table

    ' Find the table on the slide
    On Error Resume Next
    Set shape = ActiveWindow.View.slide.Shapes(tableName)
    On Error GoTo 0

    ' Check if the shape exists and is a table
    If Not shape Is Nothing Then
        If shape.Type = msoTable Then
            Set tbl = shape.table
            Debug.Print "Processing table: " & tableName

            ' Apply logic for visibleCount (example for Brand_List_2)
            If tableName = "Brand_List_2" And visibleCount = 3 Then
                ' Remove text from column 2, rows 1, 2, and 3
                tbl.cell(1, 2).shape.TextFrame.textRange.text = ""
                tbl.cell(2, 2).shape.TextFrame.textRange.text = ""
                tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
                Debug.Print "Removed text from column 2, rows 1-3 in Brand_List_2."
            End If

        Else
            Debug.Print "Shape found but is not a table: " & tableName
        End If
    Else
        Debug.Print "Table not found: " & tableName
    End If
End Sub



Sub ClearTableRows(visibleCount As Integer, tableName As String)
    Dim shape As shape
    Dim tbl As table

    ' Find the table on the slide
    On Error Resume Next
    Set shape = ActiveWindow.View.slide.Shapes(tableName)
    On Error GoTo 0

    ' If the table exists, proceed with modifications
    If Not shape Is Nothing And shape.Type = msoTable Then
        Set tbl = shape.table
        Debug.Print "Processing table: " & tableName

        Select Case visibleCount
            Case 8
                tbl.cell(3, 1).shape.Fill.visible = msoFalse ' No color
            Case 7
                tbl.cell(2, 1).shape.Fill.visible = msoFalse ' No color
                tbl.cell(3, 1).shape.Fill.visible = msoFalse ' No color
            Case 6
                tbl.cell(1, 1).shape.Fill.visible = msoFalse ' No color
                tbl.cell(2, 1).shape.Fill.visible = msoFalse ' No color
                tbl.cell(3, 1).shape.Fill.visible = msoFalse ' No color
            Case 5
                If tableName = "Brand_List_2" Then
                    tbl.cell(3, 1).shape.Fill.visible = msoFalse
                End If
                tbl.cell(1, 1).shape.Fill.visible = msoFalse
                tbl.cell(2, 1).shape.Fill.visible = msoFalse
                tbl.cell(3, 1).shape.Fill.visible = msoFalse
            Case 4
                If tableName = "Brand_List_2" Then
                    tbl.cell(2, 1).shape.Fill.visible = msoFalse
                    tbl.cell(3, 1).shape.Fill.visible = msoFalse
                End If
                tbl.cell(1, 1).shape.Fill.visible = msoFalse
                tbl.cell(2, 1).shape.Fill.visible = msoFalse
                tbl.cell(3, 1).shape.Fill.visible = msoFalse
            Case 3
                tbl.cell(1, 1).shape.Fill.visible = msoFalse
                tbl.cell(2, 1).shape.Fill.visible = msoFalse
                tbl.cell(3, 1).shape.Fill.visible = msoFalse
            Case 2
                If tableName = "Brand_List_1" Then
                    tbl.cell(3, 1).shape.Fill.visible = msoFalse
                End If
                tbl.cell(1, 1).shape.Fill.visible = msoFalse
                tbl.cell(2, 1).shape.Fill.visible = msoFalse
                tbl.cell(3, 1).shape.Fill.visible = msoFalse
        End Select

        Debug.Print "Processed table: " & tableName & " for visibleCount: " & visibleCount
    Else
        Debug.Print "Table " & tableName & " not found. Skipping."
    End If
End Sub

