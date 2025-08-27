Attribute VB_Name = "Table_Tool_6"
Sub Table_Tool_6()
    Dim selectedTable As table
    Dim shapeItem As shape
    Dim i As Integer
    Dim insertText As String

    Debug.Print "=== Running Table_Tool_6 ==="

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

    ' === Start from row 2 and number down the rows in column 1 ===
    For i = 2 To selectedTable.Rows.count
        insertText = i - 1 & "." ' Number starting from 1
        selectedTable.cell(i, 1).shape.TextFrame.textRange.text = insertText
        
        ' Right-center the text in the cell
        selectedTable.cell(i, 1).shape.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignRight
    Next i

    Debug.Print "Table_Tool_6 completed successfully."
End Sub

