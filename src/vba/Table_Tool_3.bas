Attribute VB_Name = "Table_Tool_3"
Sub Table_Tool_3()
    Dim pptSlide As slide
    Dim selectedTable As table
    Dim shapeItem As shape
    Dim insertText As String
    Dim cleanMetric As String
    Dim metricForText As String

    Debug.Print "=== Running Table_Tool_3 ==="

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
    If metric_decision = "" Then
        Debug.Print "Metric decision is empty. Exiting."
        Exit Sub
    End If
    If row_decision < 1 Then
        Debug.Print "Row decision is invalid (" & row_decision & "). Exiting."
        Exit Sub
    End If

    ' === Apply lowercase conversion based on metric_decision ===
    If metric_decision = "1" Or metric_decision = "2" Or metric_decision = "3" Then
        metricForText = LCase(metric_decision) ' Force total lowercase for 1-3
    Else
        metricForText = metric_decision ' Keep original casing for 4 and above
    End If

    Debug.Print "Metric Decision (FORCED lowercase for 1-3): " & metricForText

    ' === Generate the text to insert in Column 1, Row 1 based on metric_decision ===
    If metric_decision = "1" Or metric_decision = "2" Or metric_decision = "3" Then
        ' For metric_decision = 1-3, use "drivers of"
        insertText = "Top " & row_decision & " drivers of " & metricForText
    ElseIf metric_decision = "4" Then
        ' For metric_decision = 4, use "stated drivers"
        insertText = "Top " & row_decision & " stated drivers"
    ElseIf val(metric_decision) >= 5 Then
        ' For metric_decision >= 5, use "associations to"
        insertText = "Top " & row_decision & " associations to " & metricForText
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

    Debug.Print "Table_Tool_3 completed successfully."
End Sub

