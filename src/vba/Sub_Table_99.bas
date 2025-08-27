Attribute VB_Name = "Sub_Table_99"
Sub Sub_Table_99()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    ' === Find TARGET Table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            If tableShape.Name = "TARGET" Then
                Set targetTable = tableShape.table
                Exit For
            End If
        End If
    Next tableShape

    ' If TARGET table is not found
    If targetTable Is Nothing Then
        MsgBox "Table 'TARGET' not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Loop Through All Cells in TARGET Table and Remove Specific Text ===
    For rowIndex = 1 To targetTable.Rows.count
        For colIndex = 1 To targetTable.Columns.count
            With targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange
                If .text = "No valid data found." Or .text = "Lorem ipsum et dolor" Then
                    .text = "" ' Clear the cell
                End If
            End With
        Next colIndex
    Next rowIndex

   
End Sub

