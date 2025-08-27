Attribute VB_Name = "Lines_15"
Sub Lines_15()
    Dim pptSlide As slide
    Dim chart As chart
    Dim shape As shape
    Dim totalSeries As Integer
    Dim brandCount As Integer
    Dim i As Integer
    Dim tbl As table
    Dim brandList3 As shape
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

    ' Get the total number of series in the chart
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

    ' Check if Brand_List_3 exists, delete if it does, and recreate it if needed
    On Error Resume Next
    Set brandList3 = pptSlide.Shapes("Brand_List_3")
    On Error GoTo 0

    If Not brandList3 Is Nothing Then
        brandList3.Delete
      
    End If

    ' Now, create a new Brand_List_3 as a copy of Brand_List_2
    On Error Resume Next
    Set shape = pptSlide.Shapes("Brand_List_2")
    On Error GoTo 0

    If Not shape Is Nothing And shape.Type = msoTable Then
        ' Copy Brand_List_2 to create Brand_List_3
        shape.Copy
        pptSlide.Shapes.Paste
        Set brandList3 = pptSlide.Shapes(pptSlide.Shapes.count) ' Get the newly pasted shape (Brand_List_3)

        ' Rename the new table to "Brand_List_3"
        brandList3.Name = "Brand_List_3"

        ' Position Brand_List_3 with horizontal placement set to 29.55 cm and vertical position to 6.52 cm
        brandList3.left = 29.55 * 28.35 ' Convert cm to points (1 cm = 28.35 points)
        brandList3.Top = 6.52 * 28.35 ' Convert cm to points (1 cm = 28.35 points)

        ' Clear rows based on brand count
        If brandCount = 7 Then
            ' Clear row 2 and row 3
            Set tbl = brandList3.table
            tbl.cell(2, 1).shape.TextFrame.textRange.text = ""
            tbl.cell(2, 2).shape.TextFrame.textRange.text = ""
            tbl.cell(3, 1).shape.TextFrame.textRange.text = ""
            tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
           
        ElseIf brandCount = 8 Then
            ' Clear only row 3
            Set tbl = brandList3.table
            tbl.cell(3, 1).shape.TextFrame.textRange.text = ""
            tbl.cell(3, 2).shape.TextFrame.textRange.text = ""
            
        End If

     
    End If
End Sub

