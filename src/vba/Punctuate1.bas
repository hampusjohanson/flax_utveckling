Attribute VB_Name = "Punctuate1"
Sub Punctuate()
    Dim shp As shape
    Dim tbl As table
    Dim Sel As Selection
    Dim row As Integer
    Dim col As Integer
    
    ' Kontrollera om en tabell är markerad
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select cells in a table.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    ' Kontrollera att det är en tabell
    If shp.HasTable Then
        Set tbl = shp.table
        
        ' Loopa igenom markerade celler
        For row = 1 To tbl.Rows.count
            For col = 1 To tbl.Columns.count
                If tbl.cell(row, col).Selected Then
                    With tbl.cell(row, col).shape.TextFrame.textRange
                        ' Lägg till en punkt om den inte redan finns
                        If right(.text, 1) <> "." Then
                            .text = .text & "."
                        End If
                        ' Centrera texten
                        .ParagraphFormat.Alignment = ppAlignCenter
                        .ParagraphFormat.SpaceBefore = 0
                        .ParagraphFormat.SpaceAfter = 0
                    End With
                End If
            Next col
        Next row
        
        
    Else
        MsgBox "Selected shape does not contain a table.", vbExclamation
    End If
End Sub


