Attribute VB_Name = "ModuleObjectsClone"
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

Sub ObjectsCloneRight()
    
    Set MyDocument = Application.ActiveWindow
    
    
    Dim oldTop, OldLeft As Double
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count = 1 Then
        
        oldTop = MyDocument.Selection.ShapeRange.Top
        OldLeft = MyDocument.Selection.ShapeRange.left
        
        Set SlideShape = MyDocument.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = oldTop
            .left = OldLeft + SlideShape.width
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = MyDocument.Selection.ShapeRange.Group
        
        oldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = oldTop
            .left = OldLeft + SlideShape.width
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
    
End Sub

Sub ObjectsCloneDown()
    
    Set MyDocument = Application.ActiveWindow
    Dim oldTop, OldLeft As Double
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf MyDocument.Selection.ShapeRange.count = 1 Then
        
        oldTop = MyDocument.Selection.ShapeRange.Top
        OldLeft = MyDocument.Selection.ShapeRange.left
        
        Set SlideShape = MyDocument.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = oldTop + SlideShape.height
            .left = OldLeft
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = MyDocument.Selection.ShapeRange.Group
        
        oldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = oldTop + SlideShape.height
            .left = OldLeft
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
    
End Sub
