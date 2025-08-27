Attribute VB_Name = "Lines_11"
Sub Lines_11a()
    Dim pptSlide As slide
    Dim chart As chart
    Dim series As series
    Dim i As Integer
    Dim shape As shape
    Dim table As table
    Dim rowIndex As Integer
    Dim seriesName As String
    Dim seriesToShow As Boolean

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the table on the slide (assumed to be the first shape)
    On Error Resume Next
    Set shape = pptSlide.Shapes(1) ' Adjust this if the table is not the first shape
    On Error GoTo 0

    ' Check if the shape is a table
    If shape.Type = msoTable Then
        Set table = shape.table

        ' Loop through all rows in the table to get the series names and "Yes"/"No"
        For rowIndex = 2 To table.Rows.count ' Assuming row 1 has headers
            seriesName = table.cell(rowIndex, 1).shape.TextFrame.textRange.text ' Get the series name
            If Trim(table.cell(rowIndex, 2).shape.TextFrame.textRange.text) = "Yes" Then
                seriesToShow = True
            Else
                seriesToShow = False
            End If

            ' Find the chart and match the series
            For Each shape In pptSlide.Shapes
                If shape.Type = msoChart Then
                    Set chart = shape.chart
                    For i = 1 To chart.SeriesCollection.count
                        Set series = chart.SeriesCollection(i)
                        If series.Name = seriesName Then
                            If seriesToShow Then
                                series.Format.line.visible = msoTrue ' Show the series
                            Else
                                series.Format.line.visible = msoFalse ' Hide the series
                            End If
                        End If
                    Next i
                End If
            Next shape
        Next rowIndex

        MsgBox "Chart updated based on selected series."
    Else
        MsgBox "No table found on the slide.", vbExclamation
    End If
End Sub

