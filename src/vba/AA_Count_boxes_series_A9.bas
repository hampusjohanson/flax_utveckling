Attribute VB_Name = "AA_Count_boxes_series_A9"
Sub AA_Count_boxes_series_A9()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim shape As shape
    Dim boxesToDelete As Integer
    Dim ShapeCount As Integer
    Dim dataPointCount As Integer
    Dim boxesNeeded As Integer
    Dim leftieBox As shape
    Dim shapeName As String
    Dim cellValue As String
    Dim targetColor As Long

    ' === Set active slide ===
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Active slide set."

    ' === Locate the first chart on the slide ===
    Dim chartShape As shape
    Set chartShape = Nothing
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape ' Use the first chart found
            Exit For
        End If
    Next shape
    On Error GoTo 0

    ' === Count the data points (observations) in the first series ===
    On Error Resume Next
    If Not chartShape Is Nothing Then
        dataPointCount = chartShape.chart.SeriesCollection(1).Points.count
    Else
        dataPointCount = 0
    End If
    On Error GoTo 0

    Debug.Print "Number of data points (observations) in the first series: " & dataPointCount

    ' === Define the target fill color (#88FFC2) ===
    targetColor = RGB(136, 255, 194)

    ' === Count existing boxes with the target fill color ===
    ShapeCount = 0
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.Fill.visible = msoTrue Then
            If shape.Fill.ForeColor.RGB = targetColor Then
                ShapeCount = ShapeCount + 1
            End If
        End If
    Next shape
    On Error GoTo 0

    Debug.Print "Number of boxes with fill color #88FFC2: " & ShapeCount

    ' === Calculate "boxes_needed" or "boxes_to_be_deleted" ===
    If ShapeCount < dataPointCount Then
        boxesNeeded = dataPointCount - ShapeCount
        boxesToDelete = 0
    ElseIf ShapeCount > dataPointCount Then
        boxesToDelete = ShapeCount - dataPointCount
        boxesNeeded = 0
    End If

    ' === Remove existing BOX table if it exists ===
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.Name = "BOX" Then
            shape.Delete
            Exit For
        End If
    Next shape
    On Error GoTo 0

    ' === Create a table to display results and name it "BOX" ===
    On Error Resume Next
    Set tableShape = pptSlide.Shapes.AddTable(5, 2, 100, 100, 250, 100) ' Adjust position and size
    tableShape.Name = "BOX"

    ' Set table content
    tableShape.table.cell(1, 1).shape.TextFrame.textRange.text = "Metric"
    tableShape.table.cell(1, 2).shape.TextFrame.textRange.text = "Value"
    tableShape.table.cell(2, 1).shape.TextFrame.textRange.text = "Boxes here:"
    tableShape.table.cell(2, 2).shape.TextFrame.textRange.text = ShapeCount
    tableShape.table.cell(3, 1).shape.TextFrame.textRange.text = "Brands here:"
    tableShape.table.cell(3, 2).shape.TextFrame.textRange.text = dataPointCount
    tableShape.table.cell(4, 1).shape.TextFrame.textRange.text = "Boxes needed:"
    tableShape.table.cell(4, 2).shape.TextFrame.textRange.text = boxesNeeded
    tableShape.table.cell(5, 1).shape.TextFrame.textRange.text = "Boxes to delete:"
    tableShape.table.cell(5, 2).shape.TextFrame.textRange.text = boxesToDelete
    On Error GoTo 0

    Debug.Print "BOX table created and named 'BOX'."

    ' === Delete specified Leftie boxes in order starting from Leftie_2 ===
    For i = 2 To boxesToDelete + 1 ' Start from Leftie_2
        shapeName = "Leftie_" & i
        Set leftieBox = Nothing
        Debug.Print "Searching for shape: " & shapeName
        On Error Resume Next
        For Each shape In pptSlide.Shapes
            If shape.Name = shapeName Then
                Set leftieBox = shape
                Debug.Print "Found and deleting: " & shapeName
                Exit For
            End If
        Next shape
        On Error GoTo 0
        
        On Error Resume Next
        If Not leftieBox Is Nothing Then
            leftieBox.Delete
        Else
            Debug.Print "Shape not found: " & shapeName
        End If
        On Error GoTo 0
    Next i

    Debug.Print "Deletion process completed."
End Sub
