Attribute VB_Name = "SUB_TABLE_SCENARIO_2_1"
Sub SUB_TABLE_SCENARIO_2_1()
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
End Sub

