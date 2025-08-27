Attribute VB_Name = "Background_1_8"
Sub Background_1_8()
    ' Reference the slide and find the tables named LEFTIE and RIGHTIE
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide reference set."

    Dim pptTableLeft As table, pptTableRight As table
    Dim s As shape
    For Each s In pptSlide.Shapes
        If s.Name = "LEFTIE" And s.HasTable Then
            Set pptTableLeft = s.table
        ElseIf s.Name = "RIGHTIE" And s.HasTable Then
            Set pptTableRight = s.table
        End If
    Next s

    If pptTableLeft Is Nothing Or pptTableRight Is Nothing Then
        Debug.Print "One or both tables not found."
        MsgBox "Tables LEFTIE or RIGHTIE not found on slide.", vbExclamation
        Exit Sub
    End If

    ' Apply punctuation and right alignment to column 1 in LEFTIE and RIGHTIE from row 3
    Dim i As Long, targetTables As Variant
    targetTables = Array(pptTableLeft, pptTableRight)
    
    For Each pptTable In targetTables
        For i = 3 To pptTable.Rows.count ' Start from row 3
            With pptTable.cell(i, 1).shape.TextFrame.textRange
                ' Add period if not present
                If right(.text, 1) <> "." Then
                    .text = .text & "."
                End If
                ' Align text to the right
                .ParagraphFormat.Alignment = ppAlignRight
                .ParagraphFormat.SpaceBefore = 0
                .ParagraphFormat.SpaceAfter = 0
            End With
        Next i
    Next pptTable

    Debug.Print "Punctuation and right alignment applied to LEFTIE and RIGHTIE."
End Sub

