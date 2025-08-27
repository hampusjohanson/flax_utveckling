Attribute VB_Name = "AB_3"
Sub AB_3()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim seriesCount As Integer
    Dim targetTable As table
    Dim shape As shape

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Locate the first chart on the slide ===
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape ' Use the first chart found
            Exit For
        End If
    Next shape

    ' Check if a chart was found
    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Count the series in the chart ===
    seriesCount = chartShape.chart.SeriesCollection.count
    Debug.Print "Number of series in the chart: " & seriesCount

    ' === Locate the "TARGET" table ===
    Set targetTable = Nothing
    For Each shape In pptSlide.Shapes
        If shape.HasTable And shape.Name = "TARGET" Then
            Set targetTable = shape.table
            Exit For
        End If
    Next shape

    ' Check if the "TARGET" table was found
    If targetTable Is Nothing Then
        MsgBox "Table 'TARGET' not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Apply Fill Colors Based on Series Count ===
    If seriesCount >= 1 Then
        targetTable.cell(2, 1).shape.Fill.ForeColor.RGB = chartShape.chart.SeriesCollection(1).Format.Fill.ForeColor.RGB
    End If
    
    If seriesCount >= 2 Then
        targetTable.cell(3, 1).shape.Fill.ForeColor.RGB = chartShape.chart.SeriesCollection(2).Format.Fill.ForeColor.RGB
    End If
    
    If seriesCount >= 3 Then
        targetTable.cell(4, 1).shape.Fill.ForeColor.RGB = chartShape.chart.SeriesCollection(3).Format.Fill.ForeColor.RGB
    End If

    ' Only apply color to Row 5 if Series 4 exists
    If seriesCount >= 4 Then
        targetTable.cell(5, 1).shape.Fill.ForeColor.RGB = chartShape.chart.SeriesCollection(4).Format.Fill.ForeColor.RGB
    Else
        targetTable.cell(5, 1).shape.Fill.visible = msoFalse ' No fill if no Series 4
    End If

End Sub

