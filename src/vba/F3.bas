Attribute VB_Name = "F3"
Sub Delete_All_Notes_And_Comments()
    Dim sld As slide
    Dim i As Integer
    
    ' Loopa igenom alla bilder i presentationen
    For Each sld In ActivePresentation.Slides
        ' Ta bort anteckningar om de finns
        If sld.HasNotesPage Then
            sld.NotesPage.Shapes.Placeholders(2).TextFrame.textRange.text = ""
        End If
        
        ' Ta bort alla kommentarer
        If sld.Comments.count > 0 Then
            For i = sld.Comments.count To 1 Step -1
                sld.Comments(i).Delete
            Next i
        End If
    Next sld
    
    MsgBox "All notes and comments have been deleted."
End Sub


Sub Delete_All_Notes()
    Dim sld As slide
    
    ' Loopa igenom alla bilder i presentationen
    For Each sld In ActivePresentation.Slides
        ' Ta bort anteckningar om de finns
        If sld.HasNotesPage Then
            sld.NotesPage.Shapes.Placeholders(2).TextFrame.textRange.text = ""
        End If
    Next sld
    
    MsgBox "All notes have been deleted."
End Sub

Sub Delete_All_Comments()
    Dim sld As slide
    Dim i As Integer
    
    ' Loopa igenom alla bilder i presentationen
    For Each sld In ActivePresentation.Slides
        ' Ta bort alla kommentarer
        If sld.Comments.count > 0 Then
            For i = sld.Comments.count To 1 Step -1
                sld.Comments(i).Delete
            Next i
        End If
    Next sld
    
    MsgBox "All comments have been deleted."
End Sub


