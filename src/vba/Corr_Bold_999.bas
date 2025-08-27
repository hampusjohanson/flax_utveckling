Attribute VB_Name = "Corr_Bold_999"
Sub Corr_Bold_999()
    Dim pptSlide As slide
    Dim s As shape
    Dim tbl As table
    Dim rowIndex As Integer, colIndex As Integer
    Dim tblCount As Integer
    Dim txtBox As shape
    Dim textRange As textRange
    Dim boldText As textRange
    Dim existingBox As shape
    Dim keyShape As shape
    Dim textValue As String
    Dim hasBold As Boolean
    Dim keyArray() As String
    Dim countArray() As Integer
    Dim i As Integer, j As Integer, foundIndex As Integer
    Dim minBoldThreshold As Integer
    Dim dictSize As Integer
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    
    ' Initialize variables
    hasBold = False
    tblCount = 0
    dictSize = 0
    
    ' Iterate through all tables on the slide
    For Each s In pptSlide.Shapes
        If s.HasTable Then
            tblCount = tblCount + 1
            Set tbl = s.table
            
            ' Scan column 2 and store only BOLDED text occurrences
            For rowIndex = 1 To tbl.Rows.count
                On Error Resume Next
                textValue = Trim(tbl.cell(rowIndex, 2).shape.TextFrame.textRange.text)
                If tbl.cell(rowIndex, 2).shape.TextFrame.textRange.Font.Bold = msoTrue Then
                    hasBold = True
                    foundIndex = -1
                    
                    ' Check if textValue already exists in keyArray
                    For j = 0 To dictSize - 1
                        If keyArray(j) = textValue Then
                            foundIndex = j
                            Exit For
                        End If
                    Next j
                    
                    ' If found, increment the count, else add new entry
                    If foundIndex <> -1 Then
                        countArray(foundIndex) = countArray(foundIndex) + 1
                    Else
                        ReDim Preserve keyArray(dictSize)
                        ReDim Preserve countArray(dictSize)
                        keyArray(dictSize) = textValue
                        countArray(dictSize) = 1
                        dictSize = dictSize + 1
                    End If
                End If
                On Error GoTo 0
            Next rowIndex
        End If
    Next s
    
    ' If no bold text is found, exit
    If Not hasBold Then
        Debug.Print "No bold text found. Text box not inserted."
        Exit Sub
    End If
    
    ' Set initial minimum bold threshold to a high number
    minBoldThreshold = tblCount  ' Assume worst case (all tables have the same bold text)
    
    ' Find the lowest count among bolded texts (to determine the threshold)
    For i = 0 To dictSize - 1
        If countArray(i) < minBoldThreshold Then
            minBoldThreshold = countArray(i)
        End If
    Next i
    
    ' Delete existing "Bold_Text" textbox if it exists
    On Error Resume Next
    Set existingBox = pptSlide.Shapes("Bold_Text")
    If Not existingBox Is Nothing Then existingBox.Delete
    On Error GoTo 0
    
    ' Create new text box
    Set txtBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 400, 50)
    txtBox.Name = "Bold_Text"
    
    ' Set text content with the correct threshold
    txtBox.TextFrame.textRange.text = "Bold = strongly correlated in " & minBoldThreshold & " out of " & tblCount & " tables"
    
    ' Set font size
    txtBox.TextFrame.textRange.Font.size = 10
    
    ' Make "Bold" bold
    Set textRange = txtBox.TextFrame.textRange
    Set boldText = textRange.Characters(1, 4) ' "Bold"
    boldText.Font.Bold = msoTrue
    
    ' Align left-wise with "Key" shape if it exists
    On Error Resume Next
    Set keyShape = pptSlide.Shapes("Key")
    If Not keyShape Is Nothing Then
        txtBox.left = keyShape.left
        Debug.Print "Aligned 'Bold_Text' with 'Key' shape."
    Else
        Debug.Print "'Key' shape not found. Default position used."
    End If
    On Error GoTo 0
    
    ' Set vertical position to exactly 5.12 cm
    txtBox.Top = 5.12 * 28.35 ' Convert cm to points (1 cm ˜ 28.35 points)
    
    Debug.Print "Text box 'Bold_Text' created with message: " & txtBox.TextFrame.textRange.text
End Sub

