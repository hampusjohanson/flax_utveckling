Attribute VB_Name = "ModuleTableTranspose"
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

Sub TableTranspose()
    
    Set MyDocument = Application.ActiveWindow
    
    
    If Not (MyDocument.Selection.Type = ppSelectionShapes Or MyDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.table
                
                Set CopyTable = Application.ActiveWindow.Selection.ShapeRange.Duplicate
                
                For RowsCount = .Rows.count To 2 Step -1
                    CopyTable.table.Rows(RowsCount).Delete
                Next RowsCount
                
                For ColsCount = .Columns.count To 2 Step -1
                    CopyTable.table.Columns(ColsCount).Delete
                Next ColsCount
                
                For RowsCount = .Rows.count To 2 Step -1
                    CopyTable.table.Columns.Add
                Next RowsCount
                
                For ColsCount = .Columns.count To 2 Step -1
                    CopyTable.table.Rows.Add
                Next ColsCount
                
                CopyTable.width = Application.ActiveWindow.Selection.ShapeRange.width
                CopyTable.Top = Application.ActiveWindow.Selection.ShapeRange.Top
                CopyTable.left = Application.ActiveWindow.Selection.ShapeRange.left
                
                For RowsCount = 1 To .Rows.count
                    For ColsCount = 1 To .Columns.count
                        
                        .cell(RowsCount, ColsCount).shape.TextFrame2.textRange.Cut
                        CopyTable.table.cell(ColsCount, RowsCount).shape.TextFrame2.textRange.Paste
                        
                    Next ColsCount
                Next RowsCount
                
            End With
            
            Application.ActiveWindow.Selection.ShapeRange.Delete
            CopyTable.Select
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
