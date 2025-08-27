Attribute VB_Name = "Table_Tool_4"
Sub Table_Tool_4()
    Dim pptSlide As slide
    Dim selectedTable As table
    Dim shapeItem As shape
    Dim currentText As String
    Dim updatedText As String
    Dim i As Integer
    Dim terms As Variant

    Debug.Print "=== Running Table_Tool_4 ==="

    ' Define the terms to convert to lowercase
    terms = Array("Sales Premium", "Volume Premium", "Price Premium", "Brand Strength", "Market Share", "Customer Loyalty")

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

    ' === Get the current text from Column 1, Row 1 ===
    On Error Resume Next
    currentText = selectedTable.cell(1, 1).shape.TextFrame.textRange.text
    If Err.Number <> 0 Or currentText = "" Then
        Debug.Print "No text found in Table, Column 1, Row 1. Exiting."
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "Current Text: " & currentText

    ' === Search and replace specific terms with lowercase ===
    updatedText = currentText
    For i = LBound(terms) To UBound(terms)
        updatedText = Replace(updatedText, terms(i), LCase(terms(i))) ' Convert only specific matches to lowercase
    Next i

    ' === Insert updated text if changes were made ===
    If updatedText <> currentText Then
        selectedTable.cell(1, 1).shape.TextFrame.textRange.text = updatedText
        Debug.Print "Updated Text: " & updatedText
    Else
        Debug.Print "No changes made. Text was already correct."
    End If

    Debug.Print "Table_Tool_4 completed successfully."
End Sub

