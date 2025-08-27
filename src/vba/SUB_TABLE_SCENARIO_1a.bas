Attribute VB_Name = "SUB_TABLE_SCENARIO_1a"
Sub SUB_TABLE_SCENARIO_1a()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim bulletCountCol5 As Integer
    Dim bulletCountCol6 As Integer
    Dim bulletCountCol7 As Integer
    Dim textRange As textRange
    Dim cellText As String
    Dim i As Integer
    
    ' === Find TARGET table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            If tableShape.Name = "TARGET" Then
                Set targetTable = tableShape.table
                Exit For
            End If
        End If
    Next tableShape

    ' If TARGET table is not found
    If targetTable Is Nothing Then
        Debug.Print "Table 'TARGET' not found on the slide."
        Exit Sub
    End If

    ' === Count bullets in Column 5, 6, and 7 (Row 3 + Row 5) ===
    bulletCountCol5 = 0
    bulletCountCol6 = 0
    bulletCountCol7 = 0

    ' Loop through row 3 and 5
    For rowIndex = 3 To 5 Step 2
        ' Column 5
        Set textRange = targetTable.cell(rowIndex, 5).shape.TextFrame.textRange
        If textRange.text <> "" Then bulletCountCol5 = bulletCountCol5 + textRange.Paragraphs.count

        ' Column 6
        Set textRange = targetTable.cell(rowIndex, 6).shape.TextFrame.textRange
        If textRange.text <> "" Then bulletCountCol6 = bulletCountCol6 + textRange.Paragraphs.count

        ' Column 7
        Set textRange = targetTable.cell(rowIndex, 7).shape.TextFrame.textRange
        If textRange.text <> "" Then bulletCountCol7 = bulletCountCol7 + textRange.Paragraphs.count
    Next rowIndex

    ' === Debug Output ===
    Debug.Print "Bullet count Column 5 (Row 3+5): " & bulletCountCol5
    Debug.Print "Bullet count Column 6 (Row 3+5): " & bulletCountCol6
    Debug.Print "Bullet count Column 7 (Row 3+5): " & bulletCountCol7

    ' === Check Trigger Conditions ===
    If bulletCountCol7 = 0 And bulletCountCol5 > bulletCountCol6 Then
        Debug.Print "? SCENARIO 1 TRIGGERED!"

        ' === 1. Move bullets from Column 6 to Column 7 (Row 3 & 5) ===
        For rowIndex = 3 To 5 Step 2
            targetTable.cell(rowIndex, 7).shape.TextFrame.textRange.text = targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.text ' Move text
            targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.text = "" ' Clear column 6
        Next rowIndex

        ' === 2. Change text in Column 7, Row 1 ===
        targetTable.cell(1, 7).shape.TextFrame.textRange.text = "Strong position"
        targetTable.cell(1, 7).shape.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignCenter
        targetTable.cell(1, 7).shape.TextFrame.textRange.Font.Bold = msoTrue ' Make bold

        ' === 3. Change fill color of Column 7, Row 1 ===
        targetTable.cell(1, 7).shape.Fill.ForeColor.RGB = RGB(101, 185, 180) ' #65B9B4

        ' === 4. Merge ONLY Row 1 in Column 5 & 6 and keep only Column 5's text ===
        Dim oldText As String
        oldText = targetTable.cell(1, 5).shape.TextFrame.textRange.text ' Store original text from Column 5
        targetTable.cell(1, 5).Merge targetTable.cell(1, 6)
        targetTable.cell(1, 5).shape.TextFrame.textRange.text = oldText ' Restore only Column 5 text
        targetTable.cell(1, 5).shape.Fill.ForeColor.RGB = RGB(101, 185, 180) ' Keep fill color

        ' === 5. MOVE HALF OF THE BULLETS FROM COLUMN 5 TO COLUMN 6 ===
        For rowIndex = 3 To 5 Step 2
            Dim bulletArray() As String
            Dim firstHalf As String, secondHalf As String
            Dim midPoint As Integer

            ' Split text into array
            bulletArray = Split(targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.text, vbCrLf)

            ' If more than 4 bullets, move half to the next column
            If UBound(bulletArray) > 3 Then
                midPoint = (UBound(bulletArray) + 1) \ 2 ' Find the midpoint

                ' Extract first and second half manually
                firstHalf = ""
                secondHalf = ""

                For i = LBound(bulletArray) To midPoint - 1
                    firstHalf = firstHalf & "• " & Trim(bulletArray(i)) & vbCrLf
                Next i

                For i = midPoint To UBound(bulletArray)
                    secondHalf = secondHalf & "• " & Trim(bulletArray(i)) & vbCrLf
                Next i

                ' Apply firstHalf back to Column 5
                targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.text = firstHalf
                targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.ParagraphFormat.Bullet.visible = msoTrue
                targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.ParagraphFormat.Bullet.Character = 8226

                ' Put the second half into the next column (Column 6)
                targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.text = secondHalf
                targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.ParagraphFormat.Bullet.visible = msoTrue
                targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.ParagraphFormat.Bullet.Character = 8226
            End If
        Next rowIndex

        Debug.Print "? SCENARIO 1 EXECUTED!"
    Else
        Debug.Print "? SCENARIO 1 NOT TRIGGERED."
    End If
End Sub
