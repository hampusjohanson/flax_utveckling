Attribute VB_Name = "Lines_Count"
Sub CountBrands()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim totalSeries As Integer
    Dim brandCount As Integer
    Dim i As Integer
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

    ' Initialize the brand count (excluding the last series, and focusing on visible ones)
    brandCount = 0
    For i = 1 To totalSeries - 1 ' Exclude the last series
        Set series = chart.SeriesCollection(i)

        ' Check if the series is visible (both line and markers)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            brandCount = brandCount + 1
        End If
    Next i

    ' Show the brand count in a message box
    MsgBox "There are " & brandCount & " visible brands in the series.", vbInformation
End Sub

