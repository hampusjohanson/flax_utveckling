Attribute VB_Name = "AB_4"
Sub AB_4()
    Dim pptSlide As slide
    Dim targetTable As table
    Dim shape As shape
    Dim rowCount As Integer
    Dim i As Integer
    Dim cellText As String

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Locate the "TARGET" table ===
    Set targetTable = Nothing
    For Each shape In pptSlide.Shapes
        If shape.HasTable And shape.Name = "TARGET" Then
            Set targetTable = shape.table
            Exit For
        End If
    Next shape

    ' Check if the "TARGET" table was found
    If targetTable Is Nothing Then
        MsgBox "Table 'TARGET' not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Get row count ===
    rowCount = targetTable.Rows.count

    ' === Loop through each row in Column 2 ===
    For i = 1 To rowCount
        ' Get the text in column 2, row i
        cellText = Trim(LCase(targetTable.cell(i, 2).shape.TextFrame.textRange.text))

        ' Check if the text is "false" or "falskt" and remove it
        If cellText = "false" Or cellText = "falskt" Then
            targetTable.cell(i, 2).shape.TextFrame.textRange.text = vbNullString
        End If
    Next i

End Sub

