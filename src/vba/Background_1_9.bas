Attribute VB_Name = "Background_1_9"
Sub Background_1_9()
    ' Step 1: Reference the slide and find the tables named LEFTIE and RIGHTIE
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

    ' Step 2: Remove all fill colors (set to #F2F2F2) and set font color to 17,21,66 in columns 3 and 4 from row 3 downwards
    Dim i As Long, j As Long
    Dim targetTables As Variant
    targetTables = Array(pptTableLeft, pptTableRight)
    
    For Each pptTable In targetTables
        For i = 3 To pptTable.Rows.count
            For j = 3 To 4 ' Only column 3 and 4
                pptTable.cell(i, j).shape.Fill.BackColor.RGB = RGB(242, 242, 242) ' Set fill color to #F2F2F2
                pptTable.cell(i, j).shape.Fill.ForeColor.RGB = RGB(242, 242, 242) ' Set fill color to #F2F2F2
                pptTable.cell(i, j).shape.TextFrame.textRange.Font.color = RGB(17, 21, 66) ' 17,21,66
            Next j
        Next i
    Next pptTable

    Debug.Print "Fill colors reset to #F2F2F2 and font color set in LEFTIE and RIGHTIE."
End Sub

