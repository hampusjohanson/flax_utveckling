Attribute VB_Name = "Module27"
Sub UpdateSlideNumbersAllSlides()
    Dim sld As slide
    Dim shp As shape
    Dim SlideIndex As Integer
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        SlideIndex = sld.SlideIndex ' Get actual slide index
        
        ' Loop through all shapes in the slide
        For Each shp In sld.Shapes
            ' Check if shape is a slide number placeholder
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                    shp.TextFrame.textRange.text = SlideIndex
                End If
            End If
        Next shp
    Next sld
    
    MsgBox "Slide numbers updated for all slides!", vbInformation, "Done"
End Sub


Sub UpdateSlideNumbersExcludingHidden()
    Dim sld As slide
    Dim shp As shape
    Dim visibleSlideIndex As Integer
    visibleSlideIndex = 1  ' Start numbering from 1 (excluding hidden slides)
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        ' Check if the slide is hidden
        If sld.SlideShowTransition.Hidden = msoFalse Then
            ' Loop through all shapes in the slide
            For Each shp In sld.Shapes
                ' Check if shape is a slide number placeholder
                If shp.Type = msoPlaceholder Then
                    If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                        shp.TextFrame.textRange.text = visibleSlideIndex
                    End If
                End If
            Next shp
            visibleSlideIndex = visibleSlideIndex + 1 ' Only increment for visible slides
        Else
            ' If the slide is hidden, remove its slide number
            For Each shp In sld.Shapes
                If shp.Type = msoPlaceholder Then
                    If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                        shp.TextFrame.textRange.text = "" ' Remove number from hidden slides
                    End If
                End If
            Next shp
        End If
    Next sld
    
    MsgBox "Slide numbers updated! Hidden slides are excluded.", vbInformation, "Done"
End Sub

