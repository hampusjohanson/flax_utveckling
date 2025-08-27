Attribute VB_Name = "ModuleShapesAcrossSlides"
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

Sub DeleteTaggedShapes()
    Set MyDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For slideCount = 1 To ActivePresentation.Slides.count
                For Each shape In ActivePresentation.Slides(slideCount).Shapes
                    
                    If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        shape.Delete
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub

Sub UpdateTaggedShapePositionAndDimensions()
    Set MyDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For slideCount = 1 To ActivePresentation.Slides.count
                For Each shape In ActivePresentation.Slides(slideCount).Shapes
                    
                    If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        With shape
                            .Top = Application.ActiveWindow.Selection.ShapeRange.Top
                            .left = Application.ActiveWindow.Selection.ShapeRange.left
                            .width = Application.ActiveWindow.Selection.ShapeRange.width
                            .height = Application.ActiveWindow.Selection.ShapeRange.height
                            
                        End With
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub

Sub ShowFormCopyShapeToMultipleSlides()
    Set MyDocument = Application.ActiveWindow
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.Clear
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.columnCount = 3
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.ColumnWidths = "15;300;0"
    
    If MyDocument.Selection.Type = ppSelectionShapes Then
        
        CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.value = "NewShape" + Str(RandomNumber)
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.value = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.text = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
        End If
        
        Dim StorylineText As String
        Dim currentSlide As Long
        currentSlide = 0
        
        On Error Resume Next
        
        For slideCount = 1 To ActivePresentation.Slides.count
            
            If Not ActivePresentation.Slides(slideCount).slideNumber = Application.ActiveWindow.Selection.SlideRange.slideNumber Then
                
                StorylineText = "Untitled"
                
                On Error Resume Next
                For Each SlidePlaceHolder In ActivePresentation.Slides(slideCount).Shapes.Placeholders
                    
                    If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                        StorylineText = SlidePlaceHolder.TextFrame.textRange.text
                        Exit For
                    End If
                Next SlidePlaceHolder
                On Error GoTo 0
                
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.AddItem
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(slideCount - 1 - currentSlide, 0) = ActivePresentation.Slides(slideCount).slideNumber
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(slideCount - 1 - currentSlide, 1) = StorylineText
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(slideCount - 1 - currentSlide, 2) = ActivePresentation.Slides(slideCount).SlideID
                
            Else
                currentSlide = 1
                
            End If
            
        Next slideCount
        On Error GoTo 0
        
        CopyShapeToMultipleSlidesForm.Show
        
    Else
        MsgBox "No shapes selected."
    End If
End Sub

Sub CopyShapeToMultipleSlides()
    
    Dim shape       As shape
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim OverwriteExisting As Boolean
    Dim CrossSlideShapeId As String
    Dim SkipSlide   As Boolean
    
    OverwriteExisting = CopyShapeToMultipleSlidesForm.OptionExistingShapes1.value
    CrossSlideShapeId = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.value
    
    Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
    
    For selectedCount = 0 To CopyShapeToMultipleSlidesForm.AllSlidesListBox.ListCount - 1
        If (CopyShapeToMultipleSlidesForm.AllSlidesListBox.Selected(selectedCount) = True) Then
            
            SkipSlide = False
            
            For Each shape In ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(selectedCount))).Shapes
                
                If shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                    
                    If OverwriteExisting = True Then
                        
                        shape.Delete
                        
                    Else
                        
                        SkipSlide = True
                        
                    End If
                    
                End If
                
            Next
            
            If SkipSlide = False Then
                Application.ActiveWindow.Selection.ShapeRange.Copy
                Set pastedShape = ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(selectedCount))).Shapes.Paste
                pastedShape.Name = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.value + Str(RandomNumber)
                pastedShape.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
            End If
            
        End If
    Next selectedCount
    
    CopyShapeToMultipleSlidesForm.Hide
    Unload CopyShapeToMultipleSlidesForm
    
End Sub
