Attribute VB_Name = "AA_Count_boxes_series_1b"
Sub AA_Count_boxes_series()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim shape As shape
    Dim targetColor As Long
    Dim ShapeCount As Integer
    Dim dataPointCount As Integer
    Dim boxesNeeded As Integer
    Dim boxesToBeDeleted As Integer
    Dim leftmostBox As shape
    Dim rightmostBox As shape
    Dim minLeft As Single
    Dim maxRight As Single
    Dim tableShape As shape

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

    ' === Count the data points (observations) in the first series ===
    On Error Resume Next
    dataPointCount = chartShape.chart.SeriesCollection(1).Points.count
    On Error GoTo 0

    If dataPointCount = 0 Then
        MsgBox "Unable to determine the number of data points in the chart.", vbExclamation
        Exit Sub
    End If

    Debug.Print "Number of data points (observations) in the first series: " & dataPointCount

    ' === Define the target fill color (#88FFC2) ===
    targetColor = RGB(136, 255, 194)

    ' === Count existing boxes with the target fill color ===
    ShapeCount = 0
    minLeft = 999999 ' Initialize to a very high value
    maxRight = -1    ' Initialize to a very low value
    Set leftmostBox = Nothing
    Set rightmostBox = Nothing

    For Each shape In pptSlide.Shapes
        If shape.Fill.visible = msoTrue Then
            If shape.Fill.ForeColor.RGB = targetColor Then
                ShapeCount = ShapeCount + 1
                ' Track the leftmost box
                If shape.left < minLeft Then
                    minLeft = shape.left
                    Set leftmostBox = shape
                End If
                ' Track the rightmost box
                If (shape.left + shape.width) > maxRight Then
                    maxRight = shape.left + shape.width
                    Set rightmostBox = shape
                End If
            End If
        End If
    Next shape

    ' === Debugging Info: Print Counts to Immediate Window ===
    Debug.Print "Number of boxes with fill color #88FFC2: " & ShapeCount

    ' === Initialize the values to avoid empty variables ===
    boxesNeeded = 0
    boxesToBeDeleted = 0

    ' === Calculate "boxes_needed" or "boxes_to_be_deleted" ===
    If ShapeCount < dataPointCount Then
        boxesNeeded = dataPointCount - ShapeCount
    ElseIf ShapeCount > dataPointCount Then
        boxesToBeDeleted = ShapeCount - dataPointCount
    End If

    ' === Remove existing BOX table if it exists ===
    For Each shape In pptSlide.Shapes
        If shape.Name = "BOX" Then
            shape.Delete
            Exit For
        End If
    Next shape

    ' === Create a table to display results and name it "BOX" ===
    Set tableShape = pptSlide.Shapes.AddTable(5, 2, 100, 100, 250, 100) ' Adjust position and size
    tableShape.Name = "BOX" ' Set table name

    ' Set table title
    tableShape.table.cell(1, 1).shape.TextFrame.textRange.text = "Metric"
    tableShape.table.cell(1, 2).shape.TextFrame.textRange.text = "Value"

    ' Set table content
    tableShape.table.cell(2, 1).shape.TextFrame.textRange.text = "Boxes here:"
    tableShape.table.cell(2, 2).shape.TextFrame.textRange.text = ShapeCount

    tableShape.table.cell(3, 1).shape.TextFrame.textRange.text = "Brands here:"
    tableShape.table.cell(3, 2).shape.TextFrame.textRange.text = dataPointCount

    tableShape.table.cell(4, 1).shape.TextFrame.textRange.text = "Boxes needed:"
    tableShape.table.cell(4, 2).shape.TextFrame.textRange.text = boxesNeeded

    tableShape.table.cell(5, 1).shape.TextFrame.textRange.text = "Boxes to delete:"
    tableShape.table.cell(5, 2).shape.TextFrame.textRange.text = boxesToBeDeleted

    ' Format the table
    tableShape.table.Columns(1).width = 120
    tableShape.table.Columns(2).width = 80

    Debug.Print "BOX table created and named 'BOX'."
End Sub

