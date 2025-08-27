Attribute VB_Name = "Lines_26"
Sub Lines_26()
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

    ' Process Brand_List_1
    Call CleanTextInTable(visibleCount, "Brand_List_1")

    ' Process Brand_List_2
    Call CleanTextInTable(visibleCount, "Brand_List_2")
End Sub

Sub CleanTextInTable(visibleCount As Integer, tableName As String)
    Dim shape As shape
    Dim tbl As table

    ' Find the table on the slide
    On Error Resume Next
    Set shape = ActiveWindow.View.slide.Shapes(tableName)
    On Error GoTo 0

    ' If the table exists, apply logic for visibleCount
    If Not shape Is Nothing And shape.Type = msoTable Then
        Set tbl = shape.table
        Debug.Print "Processing table: " & tableName

        Select Case tableName
            Case "Brand_List_1"
                If visibleCount = 1 Then
                    ' Clean column 2, rows 2 and 3 (no text)
                    tbl.cell(2, 2).shape.TextFrame.textRange.text = ""
                    tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
                    Debug.Print "Cleaned column 2, rows 2-3 in Brand_List_1 for visibleCount = 1"
                ElseIf visibleCount = 2 Then
                    ' Clean column 2, row 3 (no text)
                    tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
                    Debug.Print "Cleaned column 2, row 3 in Brand_List_1 for visibleCount = 2"
                End If

            Case "Brand_List_2"
                If visibleCount = 1 Or visibleCount = 2 Then
                    ' Clean column 2, rows 1, 2, and 3 (no text)
                    tbl.cell(1, 2).shape.TextFrame.textRange.text = ""
                    tbl.cell(2, 2).shape.TextFrame.textRange.text = ""
                    tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
                    Debug.Print "Cleaned column 2, rows 1-3 in Brand_List_2 for visibleCount = " & visibleCount
                End If
        End Select

    Else
        Debug.Print "Table " & tableName & " not found. Skipping."
    End If
End Sub

