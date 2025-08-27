Attribute VB_Name = "Sub_Table_TrackBullets_Mac"
Sub Sub_Table_999()
    ' Variabler
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim bulletCount As Integer
    Dim textRange As textRange
    
    ' === Hitta TARGET-tabellen ===
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

    ' Om tabellen inte hittas
    If targetTable Is Nothing Then
        Debug.Print "Table 'TARGET' not found on the slide."
        Exit Sub
    End If

    ' === Räkna bullets i rad 3 och 5, endast kolumner 3-7 ===
    bulletCount = 0

    ' Loopa igenom rad 3 och 5
    For rowIndex = 3 To 5 Step 2 ' Endast rad 3 och 5
        For colIndex = 3 To 7 ' Endast kolumn 3 till 7
            Set textRange = targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange

            ' Om cellen inte är tom, räkna antal paragrafer (bullets)
            If textRange.text <> "" Then
                Debug.Print "Row " & rowIndex & ", Col " & colIndex & ": " & textRange.Paragraphs.count & " bullets"
                bulletCount = bulletCount + textRange.Paragraphs.count
            End If
        Next colIndex
    Next rowIndex

    ' Skriva ut totalen till Immediate Window
    Debug.Print "Total bullet points in TARGET (Rows 3 & 5, Cols 3-7): " & bulletCount
End Sub
Sub Sub_Table_TrackBullets()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim textRange As textRange
    Dim bulletCount As Integer
    Dim bulletArray(1 To 3, 1 To 5) As Integer ' Array for storing bullet counts
    
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

    ' === Count bullets and store in array ===
    Dim arrRow As Integer, arrCol As Integer
    arrRow = 1 ' Array row index
    For rowIndex = 3 To 5 Step 2 ' Rows 3 and 5
        arrCol = 1 ' Reset array column index
        For colIndex = 3 To 7 ' Columns 3 to 7
            Set textRange = targetTable.cell(rowIndex, colIndex).shape.TextFrame.textRange
            
            ' If the cell contains text, count the bullets
            If textRange.text <> "" Then
                bulletCount = textRange.Paragraphs.count
            Else
                bulletCount = 0
            End If

            ' Store in array
            bulletArray(arrRow, arrCol) = bulletCount
            
            ' Print for debugging
            Debug.Print "Row " & rowIndex & ", Col " & colIndex & " - Bullets: " & bulletCount

            arrCol = arrCol + 1 ' Move to next column in array
        Next colIndex
        arrRow = arrRow + 1 ' Move to next row in array
    Next rowIndex

    ' === Print summary from array ===
    Debug.Print "=== Bullet Count Summary ==="
    For arrRow = 1 To 3
        For arrCol = 1 To 5
            Debug.Print "Row " & (arrRow * 2 + 1) & ", Col " & (arrCol + 2) & " -> " & bulletArray(arrRow, arrCol) & " bullets"
        Next arrCol
    Next arrRow
End Sub

