Attribute VB_Name = "ModuleTableFormatting"
'MIT License

'Copyright (c) 2021 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Sub TableQuickFormat()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                TableRemoveBackgrounds
                TableRemoveBorders
                
                If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "" Then
                    TableColumnRemoveGaps
                End If
                
                If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "" Then
                    TableRowRemoveGaps
                End If
                
                For rowCount = 1 To .Rows.count
                    
                    For columnCount = 1 To .Columns.count
                        
                        .cell(rowCount, columnCount).shape.TextFrame.textRange.Font.color.RGB = RGB(0, 0, 0)

                    Next
                    
                Next
                
                For CellCount = 1 To .Rows(1).Cells.count
                    
                    .Rows(1).Cells(CellCount).Borders(ppBorderTop).Weight = 0
                    .Rows(1).Cells(CellCount).Borders(ppBorderBottom).Weight = 2
                    .Rows(1).Cells(CellCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                    .Rows(1).Cells(CellCount).shape.Fill.visible = msoFalse
                    .Rows(1).Cells(CellCount).shape.TextFrame.VerticalAnchor = msoAnchorBottom
                    .Rows(1).Cells(CellCount).shape.TextFrame.textRange.Font.Bold = msoTrue
                    .Rows(1).Cells(CellCount).shape.TextFrame.textRange.Font.color.RGB = RGB(0, 0, 0)
                    
                Next CellCount
                
                TableColumnGaps "even", 20
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRemoveBackgrounds()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.table

                .HorizBanding = False
                .VertBanding = False
                
                Application.ActiveWindow.Selection.ShapeRange.Fill.Solid
                Application.ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Application.ActiveWindow.Selection.ShapeRange.Fill.visible = msoFalse
                
                .Background.Fill.Solid
                .Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Background.Fill.visible = msoFalse
                
                ProgressForm.Show
                
                For rowCount = 1 To .Rows.count
                    
                SetProgress (rowCount / .Rows.count * 100)
                    
                    For columnCount = 1 To .Columns.count
                        
                        .cell(rowCount, columnCount).shape.Fill.Solid
                        .cell(rowCount, columnCount).shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .cell(rowCount, columnCount).shape.Fill.visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                Unload ProgressForm
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRemoveBorders()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.table
            
                ProgressForm.Show
                
                For rowCount = 1 To .Rows.count
                    
                SetProgress (rowCount / .Rows.count * 100)
                    
                    For columnCount = 1 To .Columns.count
                        
                        .cell(rowCount, columnCount).Borders(ppBorderLeft).ForeColor.RGB = RGB(255, 255, 255)
                        .cell(rowCount, columnCount).Borders(ppBorderRight).ForeColor.RGB = RGB(255, 255, 255)
                        .cell(rowCount, columnCount).Borders(ppBorderTop).ForeColor.RGB = RGB(255, 255, 255)
                        .cell(rowCount, columnCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(255, 255, 255)
                        
                        .cell(rowCount, columnCount).Borders(ppBorderLeft).Weight = 0
                        .cell(rowCount, columnCount).Borders(ppBorderRight).Weight = 0
                        .cell(rowCount, columnCount).Borders(ppBorderTop).Weight = 0
                        .cell(rowCount, columnCount).Borders(ppBorderBottom).Weight = 0
                        
                        .cell(rowCount, columnCount).Borders(ppBorderLeft).visible = msoFalse
                        .cell(rowCount, columnCount).Borders(ppBorderRight).visible = msoFalse
                        .cell(rowCount, columnCount).Borders(ppBorderTop).visible = msoFalse
                        .cell(rowCount, columnCount).Borders(ppBorderBottom).visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                Unload ProgressForm
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As rgbColor)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even" Then
                
                If MsgBox("Existing column gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                    TableColumnRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                NumberOfColumns = .Columns.count
                Dim ColumnWidthArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns + 1
                    ReDim ColumnWidthArray(0)
                    
                    For columnCount = 1 To NumberOfColumns
                        ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                        ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(columnCount).width
                        ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                        
                        If columnCount = NumberOfColumns Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = GapSize
                        End If
                        
                    Next columnCount
                    
                Else
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns - 1
                    
                    For columnCount = 1 To NumberOfColumns
                        
                        If Not columnCount = 1 Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(columnCount).width
                            ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                            
                        Else
                            ReDim ColumnWidthArray(1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(columnCount).width
                        End If
                        
                    Next columnCount
                    
                End If
                
                For columnCount = NumberOfColumns To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedColumn = .Columns.Add(columnCount)
                        
                        For CellCount = 1 To AddedColumn.Cells.count
                            AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                            AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                            AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                            
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                            AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If columnCount = NumberOfColumns Then
                            
                            Set AddedColumn = .Columns.Add
                            
                            For CellCount = 1 To AddedColumn.Cells.count
                                AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                                
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not columnCount = 1 Then
                            
                            Set AddedColumn = .Columns.Add(columnCount)
                            
                            For CellCount = 1 To AddedColumn.Cells.count
                                AddedColumn.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                                
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next columnCount
                
                For columnCount = 1 To NumberOfNewColumns
                    
                    .Columns(columnCount).width = ColumnWidthArray(columnCount - 1)
                    
                Next columnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnIncreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Dim ColumnGapSetting As Double
            ColumnGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For columnCount = 1 To .Columns.count
                    
                    If (columnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not columnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(columnCount).width = .Columns(columnCount).width + ColumnGapSetting
                    End If
                    
                Next columnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnDecreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Dim ColumnGapSetting As Double
            ColumnGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeColumnGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For columnCount = 1 To .Columns.count
                    
                    If ((columnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not columnCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Columns(columnCount).width - ColumnGapSetting) >= 0)) Then
                        .Columns(columnCount).width = .Columns(columnCount).width - ColumnGapSetting
                    End If
                    
                Next columnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableColumnRemoveGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA COLUMNGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For columnCount = .Columns.count To 1 Step -1
                    
                    If (columnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not columnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(columnCount).Delete
                    End If
                    
                Next columnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As rgbColor)
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even" Then
                
                If MsgBox("Existing row gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                    TableRowRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                NumberOfRows = .Rows.count
                Dim RowHeightArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows + 1
                    ReDim RowHeightArray(0)
                    
                    For rowCount = 1 To NumberOfRows
                        ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                        RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(rowCount).height
                        RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                        
                        If rowCount = NumberOfRows Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = GapSize
                        End If
                        
                    Next rowCount
                    
                Else
                    
                    NumberOfNewRows = NumberOfRows + NumberOfRows - 1
                    
                    For rowCount = 1 To NumberOfRows
                        
                        If Not rowCount = 1 Then
                            ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(rowCount).height
                            RowHeightArray(UBound(RowHeightArray) - 2) = GapSize
                            
                        Else
                            ReDim RowHeightArray(1)
                            RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(rowCount).height
                        End If
                        
                    Next rowCount
                    
                End If
                
                For rowCount = NumberOfRows To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedRow = .Rows.Add(rowCount)
                        
                        For CellCount = 1 To AddedRow.Cells.count
                            AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                            AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                            AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                            
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                            AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If rowCount = NumberOfRows Then
                            
                            Set AddedRow = .Rows.Add
                            
                            For CellCount = 1 To AddedRow.Cells.count
                                AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                                
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not rowCount = 1 Then
                            
                            Set AddedRow = .Rows.Add(rowCount)
                            
                            For CellCount = 1 To AddedRow.Cells.count
                                AddedRow.Cells(CellCount).shape.Fill.visible = msoFalse
                                AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.textRange.Font.size = 1
                                
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginBottom = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginLeft = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginRight = 0
                                AddedRow.Cells(CellCount).shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next rowCount
                
                For rowCount = 1 To NumberOfNewRows
                    
                    .Rows(rowCount).height = RowHeightArray(rowCount - 1)
                    
                Next rowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowIncreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Dim RowGapSetting As Double
            RowGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For rowCount = 1 To .Rows.count
                    
                    If (rowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not rowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(rowCount).height = .Rows(rowCount).height + RowGapSetting
                    End If
                    
                Next rowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowDecreaseGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Dim RowGapSetting As Double
            RowGapSetting = CDbl(GetSetting("Instrumenta", "Tables", "TableStepSizeRowGaps", "1" + GetDecimalSeperator() + "0"))
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For rowCount = 1 To .Rows.count
                    
                    If ((rowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not rowCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Rows(rowCount).height - RowGapSetting) >= 0)) Then
                        .Rows(rowCount).height = .Rows(rowCount).height - RowGapSetting
                    End If
                    
                Next rowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

Sub TableRowRemoveGaps()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA ROWGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                For rowCount = .Rows.count To 1 Step -1
                    
                    If (rowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not rowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(rowCount).Delete
                    End If
                    
                Next rowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
