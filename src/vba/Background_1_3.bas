Attribute VB_Name = "Background_1_3"
Sub Background_1_3()
    ' Read the same variables as in Lines_31, but do nothing further
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim dataArray() As String
    Dim Stronger_Last_Value As Double

    ' Dynamically get the file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:\Local\exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    ' Debug output
    Debug.Print "File path: " & filePath

    ' Check if file exists
    If Dir(filePath) = "" Then
        Debug.Print "File not found: " & filePath
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the CSV file and read Stronger_Last_Value
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        Debug.Print "Row " & currentRow & ": " & lineData
        
        ' Look for row 470 (Stronger_Last_Value)
        If currentRow = 470 Then
            dataArray = Split(lineData, ";")
            Stronger_Last_Value = CDbl(dataArray(1))
            Debug.Print "Stronger_Last_Value read: " & Stronger_Last_Value
            Exit Do
        End If
    Loop
    Close #fileNumber

    ' Validate Stronger_Last_Value
    If Stronger_Last_Value < 1 Or Stronger_Last_Value > 50 Then
        Debug.Print "Invalid Stronger_Last_Value: " & Stronger_Last_Value
        MsgBox "Invalid Stronger_Last_Value: " & Stronger_Last_Value, vbExclamation
        Exit Sub
    End If

    ' Debug print final valid value
    Debug.Print "Final Stronger_Last_Value: " & Stronger_Last_Value

    ' Adjust the number of rows in LEFTIE to match Stronger_Last_Value
    Dim pptSlide As slide
    Dim pptTable As table
    Dim targetRows As Integer
    Dim currentRows As Integer
    Dim i As Integer

    ' Reference active slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide reference set."
    
    ' Find table named LEFTIE
    Dim s As shape
    For Each s In pptSlide.Shapes
        Debug.Print "Checking shape: " & s.Name
        If s.Name = "LEFTIE" And s.HasTable Then
            Set pptTable = s.table
            Exit For
        End If
    Next s

    If pptTable Is Nothing Then
        Debug.Print "Table LEFTIE not found."
        MsgBox "Table LEFTIE not found on slide.", vbExclamation
        Exit Sub
    End If

    currentRows = pptTable.Rows.count
    targetRows = Stronger_Last_Value + 2 ' Since counting from row 3
    Debug.Print "Current rows: " & currentRows & ", Target rows: " & targetRows
    
    ' Adjust table rows
    If currentRows < targetRows Then
        Debug.Print "Adding rows..."
        For i = currentRows + 1 To targetRows
            pptTable.Rows.Add
            Debug.Print "Added row: " & i
        Next i
    ElseIf currentRows > targetRows Then
        Debug.Print "Removing rows..."
        For i = currentRows To targetRows + 1 Step -1
            pptTable.Rows(i).Delete
            Debug.Print "Deleted row: " & i
        Next i
    End If
End Sub

