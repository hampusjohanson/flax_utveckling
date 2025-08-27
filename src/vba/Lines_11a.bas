Attribute VB_Name = "Lines_11a"
Sub Lines_11_12a()
    Dim pptSlide As slide
    Dim chart As chart
    Dim series As series
    Dim i As Integer
    Dim shape As shape
    Dim tbl As table
    Dim seriesName As String
    Dim rowIndex As Integer
    Dim seriesToShow As Boolean
    Dim tableShape As shape
    Dim matchedCount As Integer

    matchedCount = 0
    Set pptSlide = ActiveWindow.View.slide

    ' Hämta tabellen med val
    On Error Resume Next
    Set tableShape = pptSlide.Shapes("Choice1")
    On Error GoTo 0

    If Not tableShape Is Nothing And tableShape.Type = msoTable Then
        Set tbl = tableShape.table

        For rowIndex = 2 To tbl.Rows.count
            seriesName = Trim(tbl.cell(rowIndex, 1).shape.TextFrame.textRange.text)
            If seriesName = "" Then GoTo ContinueLoop

            seriesToShow = (LCase(Trim(tbl.cell(rowIndex, 2).shape.TextFrame.textRange.text)) = "yes")

            ' Gå igenom alla diagram på sliden
            For Each shape In pptSlide.Shapes
                If shape.Type = msoChart Then
                    Set chart = shape.chart
                    For i = 1 To chart.SeriesCollection.count
                        Set series = chart.SeriesCollection(i)
                        If Trim(series.Name) = seriesName Then
                            If seriesToShow Then
                                series.Format.line.visible = msoTrue
                                series.MarkerStyle = msoMarkerCircle
                            Else
                                series.Format.line.visible = msoFalse
                                series.MarkerStyle = msoMarkerNone
                            End If
                            matchedCount = matchedCount + 1
                            Exit For
                        End If
                    Next i
                End If
            Next shape
ContinueLoop:
        Next rowIndex

        tableShape.Delete
       

    Else
        MsgBox "Ingen tabell med namnet 'Choice1' hittades på sliden.", vbExclamation
    End If

    ' Kör ev. efterföljande makro
    On Error Resume Next
    Application.Run "SetMarkersForAllP"
    If Err.Number <> 0 Then MsgBox "Fel vid SetMarkersForAllP: " & Err.Description
    On Error GoTo 0
End Sub

