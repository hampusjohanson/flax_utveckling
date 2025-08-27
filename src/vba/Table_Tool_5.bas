Attribute VB_Name = "Table_Tool_5"
Sub Table_Tool_5()
    Dim selectedTable As table
    Dim shapeItem As shape
    Dim insertText As String
    Dim metricForText As String
    Dim metricName As String

    Debug.Print "=== Running Table_Tool_5 ==="

    ' === Check if a table is selected ===
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Debug.Print "No table selected. Exiting."
        Exit Sub
    End If

    Set shapeItem = ActiveWindow.Selection.ShapeRange(1)
    If Not shapeItem.HasTable Then
        Debug.Print "Selected shape is not a table. Exiting."
        Exit Sub
    End If
    Set selectedTable = shapeItem.table
    Debug.Print "Table selected: " & shapeItem.Name

    ' === Exit if no valid metric or row count ===
    If number_metric < 1 Then
        Debug.Print "Metric decision is invalid (" & number_metric & "). Exiting."
        Exit Sub
    End If
    If row_decision < 1 Then
        Debug.Print "Row decision is invalid (" & row_decision & "). Exiting."
        Exit Sub
    End If

    ' === Ensure validMetrics is populated before this step ===
    If number_metric < 1 Or number_metric > validMetrics.count Then
        MsgBox "Invalid metric number: " & number_metric, vbExclamation
        Exit Sub
    End If

    ' === Get the actual metric name based on number_metric ===
    metricName = validMetrics(number_metric) ' Get the metric name from the validMetrics collection

    ' === If number_metric >= 5, remove "Absolute" from the metric name ===
    If number_metric >= 5 Then
        metricName = Replace(metricName, "Absolute", "")
    End If

    ' === Apply lowercase conversion based on number_metric ===
    If number_metric = 1 Or number_metric = 2 Or number_metric = 3 Then
        metricForText = LCase(metricName) ' Force lowercase for 1-3
    Else
        metricForText = metricName ' Keep original casing for others
    End If

    Debug.Print "Metric Name (FORCED lowercase for 1-3): " & metricForText

    ' === Generate the text to insert in Column 1, Row 1 based on number_metric ===
    If number_metric >= 1 And number_metric <= 3 Then
        ' For number_metric = 1-3, use "drivers of"
        insertText = "Top " & row_decision & " drivers of " & metricForText
    ElseIf number_metric = 4 Then
        ' For number_metric = 4, use "stated drivers"
        insertText = "Top " & row_decision & " stated drivers"
    ElseIf number_metric >= 5 Then
        ' For number_metric >= 5, use "associations to"
        insertText = "Top " & row_decision & " associations to " & metricForText
    Else
        ' If number_metric is not valid, create default text
        insertText = "Invalid metric decision"
    End If

    Debug.Print "Generated Insert Text: " & insertText

    ' === Insert into Column 1, Row 1 ===
    On Error Resume Next
    selectedTable.cell(1, 1).shape.TextFrame.textRange.text = insertText
    If Err.Number <> 0 Then
        Debug.Print "Error inserting text: " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "Table_Tool_5 completed successfully."
End Sub

