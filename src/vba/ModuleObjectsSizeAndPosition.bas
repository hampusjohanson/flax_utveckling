Attribute VB_Name = "ModuleObjectsSizeAndPosition"
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

Sub ObjectsSizeToTallest()
    Set MyDocument = Application.ActiveWindow
    Dim Tallest     As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Tallest = MyDocument.Selection.ChildShapeRange(1).height
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.height > Tallest Then Tallest = SlideShape.height
        Next
        
        MyDocument.Selection.ChildShapeRange.height = Tallest
        
    Else
        Tallest = MyDocument.Selection.ShapeRange(1).height
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.height > Tallest Then Tallest = SlideShape.height
        Next
        
        MyDocument.Selection.ShapeRange.height = Tallest
        
    End If
    
End Sub

Sub ObjectsSizeToShortest()
    Set MyDocument = Application.ActiveWindow
    Dim Shortest    As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Shortest = MyDocument.Selection.ChildShapeRange(1).height
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.height < Shortest Then Shortest = SlideShape.height
        Next
        
        MyDocument.Selection.ChildShapeRange.height = Shortest
        
    Else
        
        Shortest = MyDocument.Selection.ShapeRange(1).height
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.height < Shortest Then Shortest = SlideShape.height
        Next
        
        MyDocument.Selection.ShapeRange.height = Shortest
        
    End If
    
End Sub

Sub ObjectsSizeToWidest()
    Set MyDocument = Application.ActiveWindow
    Dim Widest      As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Widest = MyDocument.Selection.ChildShapeRange(1).width
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.width > Widest Then Widest = SlideShape.width
        Next
        
        MyDocument.Selection.ChildShapeRange.width = Widest
        
    Else
        Widest = MyDocument.Selection.ShapeRange(1).width
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.width > Widest Then Widest = SlideShape.width
        Next
        
        MyDocument.Selection.ShapeRange.width = Widest
        
    End If
    
End Sub

Sub ObjectsSizeToNarrowest()
    Set MyDocument = Application.ActiveWindow
    Dim Narrowest   As Single
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If MyDocument.Selection.HasChildShapeRange Then
        
        Narrowest = MyDocument.Selection.ChildShapeRange(1).width
        
        For Each SlideShape In MyDocument.Selection.ChildShapeRange
            If SlideShape.width < Narrowest Then Narrowest = SlideShape.width
        Next
        
        MyDocument.Selection.ChildShapeRange.width = Narrowest
        
    Else
        
        Narrowest = MyDocument.Selection.ShapeRange(1).width
        
        For Each SlideShape In MyDocument.Selection.ShapeRange
            If SlideShape.width < Narrowest Then Narrowest = SlideShape.width
        Next
        
        MyDocument.Selection.ShapeRange.width = Narrowest
        
    End If
    
End Sub

Sub ObjectsSameHeight()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.height = MyDocument.Selection.ChildShapeRange(1).height
        Else
            MyDocument.Selection.ShapeRange.height = MyDocument.Selection.ShapeRange(1).height
        End If
        
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.height = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count).height
        Else
            MyDocument.Selection.ShapeRange.height = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count).height
        End If
        
    End If
    
End Sub

Sub ObjectsSameWidth()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.width = MyDocument.Selection.ChildShapeRange(1).width
        Else
            MyDocument.Selection.ShapeRange.width = MyDocument.Selection.ShapeRange(1).width
        End If
    Else
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.width = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count).width
        Else
            MyDocument.Selection.ShapeRange.width = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count).width
        End If
    End If
End Sub

Sub ObjectsSameHeightAndWidth()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    If GetSetting("Instrumenta", "AlignDistributeSize", "DefaultTransformationMethod", "0") = 0 Then
        
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.height = MyDocument.Selection.ChildShapeRange(1).height
            MyDocument.Selection.ChildShapeRange.width = MyDocument.Selection.ChildShapeRange(1).width
            
        Else
            MyDocument.Selection.ShapeRange.height = MyDocument.Selection.ShapeRange(1).height
            MyDocument.Selection.ShapeRange.width = MyDocument.Selection.ShapeRange(1).width
        End If
        
    Else
        
        If MyDocument.Selection.HasChildShapeRange Then
            MyDocument.Selection.ChildShapeRange.height = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count).height
            MyDocument.Selection.ChildShapeRange.width = MyDocument.Selection.ChildShapeRange(MyDocument.Selection.ChildShapeRange.count).width
            
        Else
            MyDocument.Selection.ShapeRange.height = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count).height
            MyDocument.Selection.ShapeRange.width = MyDocument.Selection.ShapeRange(MyDocument.Selection.ShapeRange.count).width
        End If
        
    End If
    
End Sub

Sub ObjectsSwapPosition()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2 As Single
    
    If ActiveWindow.Selection.ShapeRange.count = 2 Then
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).left
        Left2 = ActiveWindow.Selection.ShapeRange(2).left
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        
        ActiveWindow.Selection.ShapeRange(1).left = Left2
        ActiveWindow.Selection.ShapeRange(2).left = Left1
        ActiveWindow.Selection.ShapeRange(1).Top = Top2
        ActiveWindow.Selection.ShapeRange(2).Top = Top1
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.count = 2 Then
            
            Left1 = MyDocument.Selection.ChildShapeRange(1).left
            Left2 = MyDocument.Selection.ChildShapeRange(2).left
            Top1 = MyDocument.Selection.ChildShapeRange(1).Top
            Top2 = MyDocument.Selection.ChildShapeRange(2).Top
            
            MyDocument.Selection.ChildShapeRange(1).left = Left2
            MyDocument.Selection.ChildShapeRange(2).left = Left1
            MyDocument.Selection.ChildShapeRange(1).Top = Top2
            MyDocument.Selection.ChildShapeRange(2).Top = Top1
            
        Else
            
            MsgBox "Select two shapes to swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes to swap positions."
        
    End If
    
End Sub

Sub ObjectsSwapPositionCentered()
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then Exit Sub
    
    Dim Left1, Left2, Top1, Top2, Width1, Width2, Height1, Height2 As Single
    
    If ActiveWindow.Selection.ShapeRange.count = 2 Then
        
        Left1 = ActiveWindow.Selection.ShapeRange(1).left
        Left2 = ActiveWindow.Selection.ShapeRange(2).left
        Top1 = ActiveWindow.Selection.ShapeRange(1).Top
        Top2 = ActiveWindow.Selection.ShapeRange(2).Top
        Width1 = ActiveWindow.Selection.ShapeRange(1).width
        Width2 = ActiveWindow.Selection.ShapeRange(2).width
        Height1 = ActiveWindow.Selection.ShapeRange(1).height
        Height2 = ActiveWindow.Selection.ShapeRange(2).height
        
        ActiveWindow.Selection.ShapeRange(1).left = Left2 + (Width2 - Width1) / 2
        ActiveWindow.Selection.ShapeRange(2).left = Left1 + (Width1 - Width2) / 2
        ActiveWindow.Selection.ShapeRange(1).Top = Top2 + (Height2 - Height1) / 2
        ActiveWindow.Selection.ShapeRange(2).Top = Top1 + (Height1 - Height2) / 2
        
    ElseIf MyDocument.Selection.HasChildShapeRange Then
        
        If MyDocument.Selection.ChildShapeRange.count = 2 Then
            
            Left1 = MyDocument.Selection.ChildShapeRange(1).left
            Left2 = MyDocument.Selection.ChildShapeRange(2).left
            Top1 = MyDocument.Selection.ChildShapeRange(1).Top
            Top2 = MyDocument.Selection.ChildShapeRange(2).Top
            
            Width1 = ActiveWindow.Selection.ChildShapeRange(1).width
            Width2 = ActiveWindow.Selection.ChildShapeRange(2).width
            Height1 = ActiveWindow.Selection.ChildShapeRange(1).height
            Height2 = ActiveWindow.Selection.ChildShapeRange(2).height
            
            ActiveWindow.Selection.ChildShapeRange(1).left = Left2 + (Width2 - Width1) / 2
            ActiveWindow.Selection.ChildShapeRange(2).left = Left1 + (Width1 - Width2) / 2
            ActiveWindow.Selection.ChildShapeRange(1).Top = Top2 + (Height2 - Height1) / 2
            ActiveWindow.Selection.ChildShapeRange(2).Top = Top1 + (Height1 - Height2) / 2
            
        Else
            
            MsgBox "Select two shapes to swap positions."
            
        End If
        
    Else
        
        MsgBox "Select two shapes to swap positions."
        
    End If
    
End Sub
