Attribute VB_Name = "SUB_TABLE_SCENARIO_3_Total"
Sub SUB_TABLE_SCENARIO_3_Total()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer, colIndex As Integer
    Dim textRange As textRange
    Dim rawText As String
    Dim bullets() As String
    Dim i As Integer, bulletCount As Integer
    Dim bulletCounts(3 To 7) As Integer
    Dim top_bullet_column As Integer
    Dim maxBullets As Integer
    Dim mergeText As String
    Dim fillColorCol4 As Long, fillColorCol5 As Long

    ' === Find TARGET table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing
    Set targetTable = Nothing

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
        Debug.Print "? ERROR: Table 'TARGET' not found!"
        Exit Sub
    End If

    ' === Reset bullet counts ===
    For colIndex = 3 To 7
        bulletCounts(colIndex) = 0
    Next colIndex

    ' === Loop through row 3 and row 5, count valid bullets (LEN > 3) ===
    For rowIndex = 3 To 5 Step 2
        For colIndex = 3 To 7
            Set textRange = targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange
            rawText = Trim(textRange.text)
            bulletCount = 0

            ' Split text into lines using vbLf (handling line breaks)
            If rawText <> "" Then
                bullets = Split(rawText, vbLf)
                For i = LBound(bullets) To UBound(bullets)
                    If Len(Trim(bullets(i))) > 3 Then
                        bulletCount = bulletCount + 1
                    End If
                Next i
            End If

            bulletCounts(colIndex) = bulletCounts(colIndex) + bulletCount
        Next colIndex
    Next rowIndex

    ' === Find the column with the most valid bullets ===
    maxBullets = 0
    top_bullet_column = 3 ' Default to first column checked

    For colIndex = 3 To 7
        If bulletCounts(colIndex) > maxBullets Then
            maxBullets = bulletCounts(colIndex)
            top_bullet_column = colIndex
        End If
    Next colIndex

    ' === Debug Output ===
    Debug.Print "FINAL FILTERED BULLET COUNT (LEN > 3):"
    For colIndex = 3 To 7
        Debug.Print "Column " & colIndex & " (Row 3+5): " & bulletCounts(colIndex)
    Next colIndex
    Debug.Print "Top bullet column: " & top_bullet_column

    ' === Condition Check: If Column 5 is top and Column 3 = 0, Move & Merge ===
    If top_bullet_column = 5 And bulletCounts(3) = 0 Then
        Debug.Print "? Condition Met! Moving Column 4 to 3 and merging Column 4 & 5..."

        ' === Store text and fill color from old Column 4 ===
        fillColorCol4 = targetTable.cell(1, 4).shape.Fill.ForeColor.RGB
        targetTable.cell(1, 3).shape.TextFrame.textRange.text = targetTable.cell(1, 4).shape.TextFrame.textRange.text
        targetTable.cell(3, 3).shape.TextFrame.textRange.text = targetTable.cell(3, 4).shape.TextFrame.textRange.text
        targetTable.cell(5, 3).shape.TextFrame.textRange.text = targetTable.cell(5, 4).shape.TextFrame.textRange.text

        ' Apply old Column 4 fill color to new Column 3
        targetTable.cell(1, 3).shape.Fill.ForeColor.RGB = fillColorCol4

        ' === Store text and fill color from Column 5 ===
        mergeText = targetTable.cell(1, 5).shape.TextFrame.textRange.text
        fillColorCol5 = targetTable.cell(1, 5).shape.Fill.ForeColor.RGB

        ' === Merge Column 4 & 5 in Row 1, keeping Column 5’s text ===
        targetTable.cell(1, 4).Merge targetTable.cell(1, 5)
        targetTable.cell(1, 4).shape.TextFrame.textRange.text = mergeText
        targetTable.cell(1, 4).shape.Fill.ForeColor.RGB = fillColorCol5

        ' === Move half of Column 5's bullets into the new merged Column 4 ===
        For rowIndex = 3 To 5 Step 2
            Set textRange = targetTable.cell(rowIndex, 5).shape.TextFrame.textRange
            bullets = Split(textRange.text, vbLf)
            bulletCount = UBound(bullets) + 1

            If bulletCount > 1 Then
                Debug.Print "Moving half of Column 5 bullets into Column 4."

                ' Split bullets into two parts
                midPoint = bulletCount \ 2
                firstHalf = ""
                secondHalf = ""

                For i = 0 To UBound(bullets)
                    If i < midPoint Then
                        firstHalf = firstHalf & bullets(i) & vbCrLf
                    Else
                        secondHalf = secondHalf & bullets(i) & vbCrLf
                    End If
                Next i

                ' Apply split bullets
                targetTable.cell(rowIndex, 4).shape.TextFrame.textRange.text = Trim(firstHalf)
                targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.text = Trim(secondHalf)
            End If
        Next rowIndex

        Debug.Print "? Scenario 3 Execution Complete! Column 3 inherits Column 4's color and text."
    Else
        Debug.Print "? Condition Not Met. Ending Macro Chain."
        Exit Sub
    End If
End Sub

