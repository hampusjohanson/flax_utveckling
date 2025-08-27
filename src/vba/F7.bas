Attribute VB_Name = "F7"

Sub Format_Selected_Tables_As_Leftmost()
    Dim shape As shape
    Dim sourceTable As table
    Dim targetTable As table
    Dim leftmostPos As Single
    Dim i As Long, j As Long, k As Integer
    
    On Error GoTo ErrorHandler
    
    ' Säkerställ att något är markerat
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select at least two tables.", vbExclamation
        Exit Sub
    End If
    
    ' Hitta den vänstersta tabellen i markeringen
    leftmostPos = 9999
    For Each shape In ActiveWindow.Selection.ShapeRange
        If shape.HasTable Then
            If shape.left < leftmostPos Then
                leftmostPos = shape.left
                Set sourceTable = shape.table
            End If
        End If
    Next shape
    
    ' Säkerställ att en referenstabell har hittats
    If sourceTable Is Nothing Then
        MsgBox "No valid tables selected.", vbExclamation
        Exit Sub
    End If
    
    ' Loopa igenom alla markerade tabeller och applicera formatering
    For Each shape In ActiveWindow.Selection.ShapeRange
        If shape.HasTable And Not shape.table Is sourceTable Then
            Set targetTable = shape.table
            
            ' Rensa alla befintliga kanter i mål-tabellen
            For i = 1 To targetTable.Rows.count
                For j = 1 To targetTable.Columns.count
                    For k = 1 To 4
                        targetTable.cell(i, j).Borders(k).visible = msoFalse
                    Next k
                Next j
            Next i
            
            ' Kopiera storlek, format och stil från källtabellen
            For i = 1 To sourceTable.Rows.count
                If i <= targetTable.Rows.count Then
                    targetTable.Rows(i).height = sourceTable.Rows(i).height
                End If
            Next i
            
            For j = 1 To sourceTable.Columns.count
                If j <= targetTable.Columns.count Then
                    targetTable.Columns(j).width = sourceTable.Columns(j).width
                End If
            Next j
            
            ' Kopiera cellformat
            For i = 1 To sourceTable.Rows.count
                For j = 1 To sourceTable.Columns.count
                    If i <= targetTable.Rows.count And j <= targetTable.Columns.count Then
                        ' Kopiera textstil
                        targetTable.cell(i, j).shape.TextFrame.textRange.Font.size = _
                            sourceTable.cell(i, j).shape.TextFrame.textRange.Font.size
                        
                        targetTable.cell(i, j).shape.TextFrame.textRange.Font.Name = _
                            sourceTable.cell(i, j).shape.TextFrame.textRange.Font.Name
                        
                        targetTable.cell(i, j).shape.TextFrame.textRange.ParagraphFormat.Alignment = _
                            sourceTable.cell(i, j).shape.TextFrame.textRange.ParagraphFormat.Alignment
                        
                        ' Kopiera vertikal justering
                        targetTable.cell(i, j).shape.TextFrame.VerticalAnchor = _
                            sourceTable.cell(i, j).shape.TextFrame.VerticalAnchor
                        
                        ' Kopiera marginaler
                        With sourceTable.cell(i, j).shape.TextFrame
                            targetTable.cell(i, j).shape.TextFrame.MarginBottom = .MarginBottom
                            targetTable.cell(i, j).shape.TextFrame.MarginLeft = .MarginLeft
                            targetTable.cell(i, j).shape.TextFrame.MarginRight = .MarginRight
                            targetTable.cell(i, j).shape.TextFrame.MarginTop = .MarginTop
                        End With
                        
                        ' Kopiera kanter
                        Call CopyBorders(sourceTable.cell(i, j), targetTable.cell(i, j))
                    End If
                Next j
            Next i
        End If
    Next shape
    
    MsgBox "All selected tables have been formatted to match the leftmost table.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

' Kopierar kantlinjer från en cell till en annan
Sub CopyBorders(sourceCell As cell, targetCell As cell)
    Dim borderType As Integer
    For borderType = 1 To 4 ' Top, Left, Bottom, Right borders
        With sourceCell.Borders(borderType)
            If .visible Then
                targetCell.Borders(borderType).visible = msoTrue
                targetCell.Borders(borderType).ForeColor.RGB = .ForeColor.RGB
                targetCell.Borders(borderType).Weight = .Weight
                targetCell.Borders(borderType).Style = .Style
            Else
                targetCell.Borders(borderType).visible = msoFalse
            End If
        End With
    Next borderType
End Sub

