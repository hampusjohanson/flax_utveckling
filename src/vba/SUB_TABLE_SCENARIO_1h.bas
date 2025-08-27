Attribute VB_Name = "SUB_TABLE_SCENARIO_1h"
Sub SUB_TABLE_SCENARIO_1h()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim textRange As textRange
    Dim rowIndex As Integer, colIndex As Integer
    Dim bulletCount As Integer

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

            ' If there is more than one paragraph, simulate backspace to remove last empty one
            If bulletCount > 1 Then
                textRange.Characters(Len(textRange.text)).Delete ' Simulate Backspace
                Debug.Print "? Simulated Backspace in Row " & rowIndex & ", Column " & colIndex
            End If
        Next colIndex
    Next rowIndex

    Debug.Print "? SUB_TABLE_SCENARIO_1h EXECUTED! Empty bullet rows removed using Backspace."
End Sub

