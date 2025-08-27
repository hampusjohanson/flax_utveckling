Attribute VB_Name = "SUB_TABLE_SCENARIO_2_2"
Sub SUB_TABLE_SCENARIO_2_2()
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
    Dim mergeText As String
    Dim fillColor As Long

    ' === Find TARGET table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing
    Set targetTable = Nothing

    ' Debugging: Print all shape names
    Debug.Print "Checking shapes on slide..."
    
    For Each tableShape In pptSlide.Shapes
        Debug.Print "Shape: " & tableShape.Name & " | HasTable: " & tableShape.HasTable
        If tableShape.Name = "TARGET" And tableShape.HasTable Then
            Set targetTable = tableShape.table
            Debug.Print "? Found Table 'TARGET'!"
            Exit For
        End If
    Next tableShape

    ' If TARGET table is not found
    If targetTable Is Nothing Then
        Debug.Print "? ERROR: Table 'TARGET' was found but could not be assigned in VBA!"
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

    ' === Debug Output ===
    Debug.Print "FINAL FILTERED BULLET COUNT (LEN > 3):"
    For colIndex = 3 To 7
        Debug.Print "Column " & colIndex & " (Row 3+5): " & bulletCounts(colIndex)
    Next colIndex

    ' === Merge Row 1, Column 3 and 4 (Always) ===
    Debug.Print "? Merging Row 1, Column 3 and 4 regardless of bullet counts..."
        
    ' Store text from Column 4, Row 1
    mergeText = targetTable.cell(1, 4).shape.TextFrame.textRange.text

    ' Store fill color from Column 4, Row 1
    fillColor = targetTable.cell(1, 4).shape.Fill.ForeColor.RGB

    ' Merge Column 3 and 4 in Row 1
    targetTable.cell(1, 3).Merge targetTable.cell(1, 4)

    ' Set the merged cell's text
    targetTable.cell(1, 3).shape.TextFrame.textRange.text = mergeText

    ' Apply the old fill color from Column 4
    targetTable.cell(1, 3).shape.Fill.ForeColor.RGB = fillColor

    Debug.Print "? Merge Complete! Fill color preserved."
    
End Sub

