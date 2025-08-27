Attribute VB_Name = "Corr_Bold_5"
Sub Corr_Bold_5()
    Dim pptSlide As slide
    Dim s As shape
    Dim rowIndex As Integer
    Dim tbl As table
    Dim tblCount As Integer
    Dim tbls As Collection
    Dim textValue As String
    Dim occurrences() As Variant
    Dim found As Boolean
    Dim i As Integer
    Dim uniqueCount As Integer
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    
    ' Initialize collection and variables
    Set tbls = New Collection
    tblCount = 0
    uniqueCount = 0

    ' Initialize occurrences array (2D array: column 1 for text, column 2 for counts)
    ReDim occurrences(1 To 48, 1 To 2) ' Up to 48 unique values

    ' Find all tables on the slide
    For Each s In pptSlide.Shapes
        If s.HasTable Then
            tbls.Add s.table
            tblCount = tblCount + 1
        End If
    Next s
    
    ' If fewer than 5 tables are found, exit
    If tblCount < 5 Then
        MsgBox "Not enough tables found on the slide (minimum 5 required).", vbExclamation
        Exit Sub
    End If

    Debug.Print "Total tables found: " & tblCount
    
    ' Scan column 2 of each table and store occurrences in the array
    For Each tbl In tbls
        For rowIndex = 1 To tbl.Rows.count
            On Error Resume Next
            textValue = Trim(tbl.cell(rowIndex, 2).shape.TextFrame.textRange.text)
            On Error GoTo 0
            
            If textValue <> "" Then
                found = False
                ' Check if textValue already exists in the array
                For i = 1 To uniqueCount
                    If occurrences(i, 1) = textValue Then
                        ' If found, increment the occurrence value
                        occurrences(i, 2) = occurrences(i, 2) + 1
                        found = True
                        Exit For
                    End If
                Next i
                
                ' If textValue is not found, add it to the array with a count of 1
                If Not found Then
                    uniqueCount = uniqueCount + 1
                    occurrences(uniqueCount, 1) = textValue
                    occurrences(uniqueCount, 2) = 1 ' Initial count
                End If
            End If
        Next rowIndex
    Next tbl

    ' Apply bold formatting if text appears in at least 5 out of ALL tables
    For Each tbl In tbls
        For rowIndex = 1 To tbl.Rows.count
            On Error Resume Next
            textValue = Trim(tbl.cell(rowIndex, 2).shape.TextFrame.textRange.text)
            On Error GoTo 0
            
            If textValue <> "" Then
                ' Check if the text has at least 5 occurrences across all tables
                For i = 1 To uniqueCount
                    If occurrences(i, 1) = textValue Then
                        If occurrences(i, 2) >= 5 Then
                            tbl.cell(rowIndex, 2).shape.TextFrame.textRange.Font.Bold = msoTrue
                            tbl.cell(rowIndex, 1).shape.TextFrame.textRange.Font.Bold = msoTrue
                        End If
                        Exit For
                    End If
                Next i
            End If
        Next rowIndex
    Next tbl

    Debug.Print "Bold formatting applied where needed."
End Sub

