Attribute VB_Name = "SUB_TABLE_SCENARIO_1g"
Sub SUB_TABLE_SCENARIO_1g()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim textRange As textRange
    Dim rowIndex As Integer, colIndex As Integer
    Dim bulletCount As Integer
    Dim cleanedText As String
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

    ' === Loop through Row 3 & 5, Columns 3 to 7 ===
    For rowIndex = 3 To 5 Step 2
        For colIndex = 3 To 7
            Set textRange = targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange

            ' Count bullets using paragraphs
            bulletCount = textRange.Paragraphs.count
            cleanedText = ""

            ' Rebuild text without empty paragraphs
            For i = 1 To bulletCount
                If Len(Trim(textRange.Paragraphs(i).text)) > 0 Then ' Ignore empty bullets
                    cleanedText = cleanedText & textRange.Paragraphs(i).text & vbCrLf
                End If
            Next i

            ' Apply cleaned text back to the cell (remove empty bullet rows)
            If Len(Trim(cleanedText)) > 0 Then
                targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = cleanedText
            Else
                targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = "" ' Ensure fully empty cells are cleared
            End If

            ' === Manually Adjust Row Height to Reduce Gaps ===
            If targetTable.Rows(rowIndex).height > 20 Then ' Ensure we don't shrink too much
                targetTable.Rows(rowIndex).height = targetTable.Rows(rowIndex).height - 5
            End If

            Debug.Print "? Cleaned Row " & rowIndex & ", Column " & colIndex
        Next colIndex
    Next rowIndex

    Debug.Print "? SUB_TABLE_SCENARIO_1g EXECUTED! Empty bullet rows removed."
End Sub

