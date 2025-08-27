Attribute VB_Name = "ModuleSpellcheckLanguage"
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


Sub ShowChangeSpellCheckLanguageForm()
    
    Dim LanguageNames(1 To 216) As String
    LanguageNames(3) = "English UK"
    LanguageNames(2) = "English US"
    LanguageNames(4) = "Norwegian Bokmol"
    LanguageNames(5) = "Norwegian Nynorsk"
    LanguageNames(1) = "Swedish"
    
    ChangeSpellCheckLanguageForm.ComboBox1.Clear
    For i = 1 To 216
        ChangeSpellCheckLanguageForm.ComboBox1.AddItem LanguageNames(i)
    Next
    
    ChangeSpellCheckLanguageForm.Show
    
End Sub

Sub ChangeSpellCheckLanguage()
    
    Dim LanguageNames(1 To 216) As String
    LanguageNames(3) = "English UK"
    LanguageNames(2) = "English US"
    LanguageNames(4) = "Norwegian Bokmol"
    LanguageNames(5) = "Norwegian Nynorsk"
    LanguageNames(1) = "Swedish"
    
    Dim LanguageIDs(1 To 216) As String
    LanguageIDs(3) = msoLanguageIDEnglishUK
    LanguageIDs(2) = msoLanguageIDEnglishUS
    LanguageIDs(4) = msoLanguageIDNorwegianBokmol
    LanguageIDs(5) = msoLanguageIDNorwegianNynorsk
    LanguageIDs(1) = msoLanguageIDSwedish
    
    ChangeSpellCheckLanguageForm.Hide
    
    Dim TargetLanguageID As String
    TargetLanguageID = LanguageIDs(ChangeSpellCheckLanguageForm.ComboBox1.ListIndex + 1)
    
    Dim TargetLanguage As String
    TargetLanguage = LanguageNames(ChangeSpellCheckLanguageForm.ComboBox1.ListIndex + 1)
    
    Dim PresentationSlide As PowerPoint.slide
    Dim SlideShape  As PowerPoint.shape
    Dim SlideSmartArtNode As SmartArtNode
    Dim GroupCount  As Integer
    
    
    #If Mac Then
    'Mac does not (yet) support property .HasHandoutMaster
        
    On Error Resume Next
    For Each SlideShape In ActivePresentation.HandoutMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
        End If
    Next
    On Error GoTo 0

    #Else
    
    If ActivePresentation.HasHandoutMaster Then
    For Each SlideShape In ActivePresentation.HandoutMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #End If
               
    If ActivePresentation.HasTitleMaster Then
    For Each SlideShape In ActivePresentation.TitleMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #If Mac Then
    'Mac does not (yet) support property .HasNotesMaster
        
    On Error Resume Next
    For Each SlideShape In ActivePresentation.NotesMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
        End If
    Next
    On Error GoTo 0
        
    #Else
    
    If ActivePresentation.HasNotesMaster Then
    For Each SlideShape In ActivePresentation.NotesMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
        End If
    Next
    End If
    
    #End If
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
    
    SetProgress (PresentationSlide.slideNumber / ActivePresentation.Slides.count * 100)
    
        For Each SlideShape In PresentationSlide.Shapes
            ChangeShapeSpellCheckLanguage SlideShape, TargetLanguageID
        Next SlideShape
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    For Each SlideShape In ActivePresentation.SlideMaster.Shapes
        ChangeShapeSpellCheckLanguage SlideShape, TargetLanguageID
    Next
    
    MsgBox "Changed spellcheck language to " + TargetLanguage + " on all slides."
    
    Unload ChangeSpellCheckLanguageForm
    
End Sub

Sub ChangeShapeSpellCheckLanguage(SlideShape, TargetLanguageID)
    
    If SlideShape.Type = msoGroup Then
        
        Set SlideShapeGroup = SlideShape.GroupItems
        
        For Each SlideShapeChild In SlideShapeGroup
            ChangeShapeSpellCheckLanguage SlideShapeChild, TargetLanguageID
        Next
        
    Else
        
        If SlideShape.HasTextFrame Then
            
            SlideShape.TextFrame2.textRange.LanguageID = TargetLanguageID
                       
        End If
        
        If SlideShape.HasTable Then
            For tableRow = 1 To SlideShape.table.Rows.count
                    For TableColumn = 1 To SlideShape.table.Columns.count
                        SlideShape.table.cell(tableRow, TableColumn).shape.TextFrame2.textRange.LanguageID = TargetLanguageID
                    Next
            Next
        End If
        
        If SlideShape.HasSmartArt Then
            
            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.count
                
                For Each SlideSmartArtNode In SlideShape.SmartArt.AllNodes
                    SlideSmartArtNode.TextFrame2.textRange.LanguageID = TargetLanguageID
                 Next
                            
            Next
            
        End If
        
    End If
    
End Sub

