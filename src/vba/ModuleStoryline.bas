Attribute VB_Name = "ModuleStoryline"
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

Sub CopySlideNotesToClipboard(ExportToWord As Boolean)
    
    Dim PresentationSlide As PowerPoint.slide
    Dim SlidePlaceHolder As PowerPoint.shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, width:=100, height:=100)
    Dim PlaceHolderTextRange As textRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.textRange
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.slideNumber / ActivePresentation.Slides.count * 100)
        
        If PresentationSlide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
            
            PresentationSlide.NotesPage.Shapes.Placeholders(2).TextFrame.textRange.Copy
            
            PlaceHolderTextRange.Characters(0).InsertAfter Chr(13) & Chr(13) & "[Slide " & Str(PresentationSlide.slideNumber) & "]" & Chr(13)
            PlaceHolderTextRange.Characters(0).Paste
            
        End If
        
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    If Not SlidePlaceHolder.TextFrame.textRange.text = "" Then
        
        SlidePlaceHolder.TextFrame.textRange.Copy
        
        If ExportToWord = True Then
            
            #If Mac Then
                
                If CheckIfAppleScriptPluginIsInstalled > 0 Then
                    Dim PasteIntoWord As String
                    PasteIntoWord = AppleScriptTask("InstrumentaAppleScriptPlugin.applescript", "PasteTextIntoWord", "")
                Else
                    MsgBox "Cannot launch Word, optional Instrumenta AppleScript not found. Slide notes are copied to clipboard."
                End If
                
            #Else
                
                Dim WordApplication, WordDocument As Object
                
                On Error Resume Next
                Set WordApplication = GetObject(Class:="Word.Application")
                Err.Clear
                
                If WordApplication Is Nothing Then Set WordApplication = CreateObject(Class:="Word.Application")
                On Error GoTo 0
                
                WordApplication.visible = True
                Set WordDocument = WordApplication.Documents.Add
                
                With WordApplication
                    .Selection.PasteAndFormat wdPasteDefault
                End With
                
            #End If
            SlidePlaceHolder.Delete
            
        Else
            SlidePlaceHolder.Delete
            MsgBox "Slide notes copied to clipboard."
        End If
        
    Else
        SlidePlaceHolder.Delete
        MsgBox "Slide notes are empty."
    End If
    
End Sub

Sub CopyStorylineToClipboard(ExportToWord As Boolean)
    
    Dim PresentationSlide As PowerPoint.slide
    Dim SlidePlaceHolder As PowerPoint.shape
    Dim ClipboardObject As Object
    Dim StorylineText As String
    
    ProgressForm.Show
    
    For Each PresentationSlide In ActivePresentation.Slides
        
        SetProgress (PresentationSlide.slideNumber / ActivePresentation.Slides.count * 100)
        
        For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
            
            If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                StorylineText = StorylineText & SlidePlaceHolder.TextFrame.textRange.text & Chr(13)
                Exit For
            End If
        Next SlidePlaceHolder
    Next PresentationSlide
    
    ProgressForm.Hide
    Unload ProgressForm
    
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, left:=0, Top:=0, width:=100, height:=100)
    SlidePlaceHolder.TextFrame.textRange.text = StorylineText
    SlidePlaceHolder.TextFrame.textRange.Copy
    SlidePlaceHolder.Delete
    
    If Not StorylineText = "" Then
        If ExportToWord = True Then
            
            #If Mac Then
            
                If CheckIfAppleScriptPluginIsInstalled > 0 Then
                    Dim PasteIntoWord As String
                    PasteIntoWord = AppleScriptTask("InstrumentaAppleScriptPlugin.applescript", "PasteTextIntoWord", "")
                Else
                    MsgBox "Cannot launch Word, optional Instrumenta AppleScript not found. Storyline is copied to clipboard."
                End If
            
            #Else
                
                Dim WordApplication, WordDocument As Object
                
                On Error Resume Next
                Set WordApplication = GetObject(Class:="Word.Application")
                Err.Clear
                
                If WordApplication Is Nothing Then Set WordApplication = CreateObject(Class:="Word.Application")
                On Error GoTo 0
                
                WordApplication.visible = True
                Set WordDocument = WordApplication.Documents.Add
                
                With WordApplication
                    
                    .Selection.PasteAndFormat wdPasteDefault
                    
                End With
                
            #End If
            
        Else
            MsgBox "Storyline copied to clipboard."
        End If
        
    Else
        MsgBox "Storyline is empty."
    End If
    
End Sub

Sub PasteStorylineInSelectedShape()
    
    Set MyDocument = Application.ActiveWindow
    If Not MyDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "Please select a shape."
    Else
        
        Dim PresentationSlide As PowerPoint.slide
        Dim SlidePlaceHolder As PowerPoint.shape
        Dim ClipboardObject As Object
        Dim StorylineText As String
        
        ProgressForm.Show
        
        For Each PresentationSlide In ActivePresentation.Slides
            
            SetProgress (PresentationSlide.slideNumber / ActivePresentation.Slides.count * 100)
            
            For Each SlidePlaceHolder In PresentationSlide.Shapes.Placeholders
                
                If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                    StorylineText = StorylineText & SlidePlaceHolder.TextFrame.textRange.text & Chr(13)
                    Exit For
                End If
            Next SlidePlaceHolder
        Next PresentationSlide
        
        ProgressForm.Hide
        Unload ProgressForm
        
        Application.ActiveWindow.Selection.ShapeRange(1).TextFrame.textRange.text = StorylineText
        
    End If
    
End Sub
