Attribute VB_Name = "Delete_4"
Sub Delete_4()
    Dim slide As slide
    Dim shape As shape
    Dim response As VbMsgBoxResult
    Dim initialSlideIndex As Integer

    ' Store the initial slide index
    initialSlideIndex = ActiveWindow.View.slide.SlideIndex

    ' Define slide dimensions
    Dim slideWidth As Single
    Dim slideHeight As Single

    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight

    ' Loop through all slides
    For Each slide In ActivePresentation.Slides
        ' Activate the slide to ensure shapes can be selected
        slide.Select
        
        ' Loop through all shapes on the slide
        For Each shape In slide.Shapes
            ' Check if the shape is partially or fully outside the slide boundaries
            If shape.left < 0 Or shape.Top < 0 Or _
               (shape.left + shape.width) > slideWidth Or _
               (shape.Top + shape.height) > slideHeight Then

             ' Highlight the shape by referencing it (no actual selection)
If response = vbYes Then
    shape.Delete
End If

                ' Ask the user whether to delete the shape
                response = MsgBox("Delete this object on Slide " & slide.SlideIndex & "?", vbYesNo + vbQuestion, "Object Outside Margin")

                If response = vbYes Then
                    shape.Delete
                End If
            End If
        Next shape
    Next slide

    ' Return to the initial slide
    ActivePresentation.Slides(initialSlideIndex).Select

    MsgBox "Finished checking all slides.", vbInformation, "Done"
End Sub


