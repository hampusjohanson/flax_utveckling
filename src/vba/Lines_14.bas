Attribute VB_Name = "Lines_14"
Sub Lines_14()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim i As Integer
    Dim tbl As table
    Dim seriesName As String
    Dim series As series
    Dim visibleCount As Integer
    Dim brandCount As Integer

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

    ' Initialize visible brand count (excluding the last series)
    visibleCount = 0
    brandCount = 0

    ' Count visible brands (excluding the last series)
    For i = 1 To chart.SeriesCollection.count - 1 ' Exclude the last series
        Set series = chart.SeriesCollection(i)

        ' Check if the series is visible (both line and markers)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleCount = visibleCount + 1
        End If
    Next i

    ' If there are 3 or fewer visible brands, display a message and exit
    If visibleCount <= 3 Then
       
        Exit Sub
    End If

    ' Find the table named "Brand_List_2"
    On Error Resume Next
    Set shape = pptSlide.Shapes("Brand_List_2")
    On Error GoTo 0

    ' If the table is found, populate it
    If Not shape Is Nothing And shape.Type = msoTable Then
        Set tbl = shape.table

        ' Initialize visibleCount for series
        visibleCount = 0

        ' Loop through the series to find the 4th to 6th visible series
        For i = 1 To chart.SeriesCollection.count
            Set series = chart.SeriesCollection(i)

            ' Check if the series is visible (both line and markers)
            If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
                visibleCount = visibleCount + 1
                seriesName = series.Name

                ' Paste series name into the table depending on the visible count
                If visibleCount = 4 Then
                    tbl.cell(1, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 1, column 2 (for 4th visible)
                ElseIf visibleCount = 5 Then
                    tbl.cell(2, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 2, column 2 (for 5th visible)
                ElseIf visibleCount = 6 Then
                    tbl.cell(3, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 3, column 2 (for 6th visible)
                End If
            End If

            ' Stop once the 6th visible series has been found
            If visibleCount = 6 Then Exit For
        Next i

        ' Clear text based on the number of visible brands
        If visibleCount = 4 Then
            tbl.cell(2, 2).shape.TextFrame.textRange.text = "" ' Clear row 2, column 2
            tbl.cell(3, 2).shape.TextFrame.textRange.text = "" ' Clear row 3, column 2
        ElseIf visibleCount = 5 Then
            tbl.cell(3, 2).shape.TextFrame.textRange.text = "" ' Clear row 3, column 2
        End If

        ' Check if at least 3 visible series were found
        If visibleCount < 3 Then
            
        End If
    Else
        
    End If
End Sub

