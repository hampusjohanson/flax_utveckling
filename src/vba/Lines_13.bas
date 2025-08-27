Attribute VB_Name = "Lines_13"
Sub Lines_13()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim i As Integer
    Dim tbl As table
    Dim seriesName As String
    Dim series As series
    Dim visibleCount As Integer

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

    ' Find the table named "Brand_List_1"
    On Error Resume Next
    Set shape = pptSlide.Shapes("Brand_List_1")
    On Error GoTo 0

    ' If the table is found, populate it
    If Not shape Is Nothing And shape.Type = msoTable Then
        Set tbl = shape.table
        visibleCount = 0

        ' Loop through the series to find the first three visible series
        For i = 1 To chart.SeriesCollection.count
            Set series = chart.SeriesCollection(i)

            ' Check if the series is visible (both line and markers)
            If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
                visibleCount = visibleCount + 1
                seriesName = series.Name

                ' Paste series name into the table depending on the visible count
                If visibleCount = 1 Then
                    tbl.cell(1, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 1, column 2
                ElseIf visibleCount = 2 Then
                    tbl.cell(2, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 2, column 2
                ElseIf visibleCount = 3 Then
                    tbl.cell(3, 2).shape.TextFrame.textRange.text = seriesName ' Paste in row 3, column 2
                End If
            End If

            ' Stop once the first 3 visible series have been found
            If visibleCount = 3 Then Exit For
        Next i

        ' Check if at least 3 visible series were found
        If visibleCount < 3 Then
            
        End If
    Else
        
    End If
End Sub

