Attribute VB_Name = "ModuleCopyPosition"
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

Public TopToCopy, LeftToCopy, WidthToCopy, HeightToCopy As Long
Public PositionCopied As Boolean

Sub CopyPosition()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    TopToCopy = MyDocument.Selection.ShapeRange(1).Top
    LeftToCopy = MyDocument.Selection.ShapeRange(1).left
    WidthToCopy = MyDocument.Selection.ShapeRange(1).width
    HeightToCopy = MyDocument.Selection.ShapeRange(1).height
    PositionCopied = True
    
    End If
End Sub

Sub PastePosition()
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    If PositionCopied = True Then
        MyDocument.Selection.ShapeRange(1).Top = TopToCopy
        MyDocument.Selection.ShapeRange(1).left = LeftToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
    End If
End Sub

Sub PastePositionAndDimensions()
    
    Set MyDocument = Application.ActiveWindow
    
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    If PositionCopied = True Then
        MyDocument.Selection.ShapeRange(1).Top = TopToCopy
        MyDocument.Selection.ShapeRange(1).left = LeftToCopy
        MyDocument.Selection.ShapeRange(1).width = WidthToCopy
        MyDocument.Selection.ShapeRange(1).height = HeightToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
    End If
End Sub
