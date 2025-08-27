Attribute VB_Name = "SUB_TABLE_SCENARIO_1d"
Sub SUB_TABLE_SCENARIO_1d()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim textRange As textRange
    Dim rowIndex As Integer
    Dim bulletCount As Integer
    Dim midPoint As Integer
    Dim firstHalf As String, secondHalf As String
    Dim i As Integer
    Dim bullets() As String

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

    ' === Process Row 3 and Row 5 in Column 5 ===
    For rowIndex = 3 To 5 Step 2
        Set textRange = targetTable.cell(rowIndex, 5).shape.TextFrame.textRange

        ' Debugging: Print raw text content
        Debug.Print "Row " & rowIndex & ", Column 5 raw text: " & vbCrLf & textRange.text

        ' Split text into bullets based on line breaks (handling vbLf for PowerPoint)
        bullets = Split(textRange.text, vbLf)
        bulletCount = UBound(bullets) + 1

        ' Debugging: Show bullet count
        Debug.Print "Row " & rowIndex & ", Column 5 has " & bulletCount & " bullets."

        ' If more than 4 bullets, split them into two halves
        If bulletCount > 4 Then
            Debug.Print "Moving half of them to Column 6."

            ' Find midpoint
            midPoint = bulletCount \ 2

            ' Extract first and second halves as strings
            firstHalf = ""
            secondHalf = ""

            For i = 0 To UBound(bullets)
                If i < midPoint Then
                    firstHalf = firstHalf & bullets(i) & vbCrLf
                Else
                    secondHalf = secondHalf & bullets(i) & vbCrLf
                End If
            Next i

            ' === Apply firstHalf back to Column 5 ===
            targetTable.cell(rowIndex, 5).shape.TextFrame.textRange.text = Trim(firstHalf)

            ' === Apply secondHalf to Column 6 ===
            targetTable.cell(rowIndex, 6).shape.TextFrame.textRange.text = Trim(secondHalf)
        Else
            Debug.Print "Row " & rowIndex & ", Column 5 has " & bulletCount & " bullets, no movement needed."
        End If
    Next rowIndex

    Debug.Print "? SUB_TABLE_SCENARIO_1d EXECUTED! Bullets moved without formatting changes."
End Sub

