Attribute VB_Name = "Lines_17"
Sub Lines_17()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim totalSeries As Integer
    Dim visibleBrandCount As Integer
    Dim i As Integer
    Dim series As series
    Dim brandList3 As shape

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

    ' Get the total number of series in the chart
    totalSeries = chart.SeriesCollection.count

    ' Initialize visible series count
    visibleBrandCount = 0

    ' Loop through all series and count visible ones
    For i = 1 To totalSeries
        Set series = chart.SeriesCollection(i)
        If series.Format.line.visible = msoTrue And series.MarkerStyle <> msoMarkerNone Then
            visibleBrandCount = visibleBrandCount + 1
        End If
    Next i

    ' Debug output for visible brands
    Debug.Print "Visible brands: " & visibleBrandCount

    ' Check if visible brands are fewer than 7
    If visibleBrandCount < 7 Then
        ' Find and delete Brand_List_3 if it exists
        On Error Resume Next
        Set brandList3 = pptSlide.Shapes("Brand_List_3")
        On Error GoTo 0

        If Not brandList3 Is Nothing Then
            brandList3.Delete
          
        Else
        
        End If
    Else
        
    End If
End Sub

