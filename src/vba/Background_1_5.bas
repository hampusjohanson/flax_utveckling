Attribute VB_Name = "Background_1_5"
Sub Background_1_5()
    ' Step 1: Read Stronger_Last_Value from row 470 and Associations from row 469
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim dataArray() As String
    Dim Stronger_Last_Value As Double
    Dim Associations_Total As Double

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

    ' Open the CSV file and read Stronger_Last_Value and Associations_Total
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        
        If currentRow = 469 Then ' Associations row
            dataArray = Split(lineData, ";")
            Associations_Total = CDbl(dataArray(1))
            Debug.Print "Associations_Total read: " & Associations_Total
        End If
        
        If currentRow = 470 Then ' Stronger_Last_Value row
            dataArray = Split(lineData, ";")
            Stronger_Last_Value = CDbl(dataArray(1))
            Debug.Print "Stronger_Last_Value read: " & Stronger_Last_Value
            Exit Do
        End If
    Loop
    Close #fileNumber

    ' Validate values
    If Stronger_Last_Value < 1 Or Stronger_Last_Value > 50 Then
        Debug.Print "Invalid Stronger_Last_Value: " & Stronger_Last_Value
        MsgBox "Invalid Stronger_Last_Value: " & Stronger_Last_Value, vbExclamation
        Exit Sub
    End If

    If Associations_Total < Stronger_Last_Value + 1 Then
        Debug.Print "Invalid Associations_Total: " & Associations_Total
        MsgBox "Invalid Associations_Total: " & Associations_Total, vbExclamation
        Exit Sub
    End If

    Debug.Print "Final Stronger_Last_Value: " & Stronger_Last_Value
    Debug.Print "Final Associations_Total: " & Associations_Total

    ' Step 2: Reference the slide and find the table named LEFTIE
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide reference set."

    Dim pptTable As table
    Dim s As shape
    For Each s In pptSlide.Shapes
        Debug.Print "Checking shape: " & s.Name
        If s.Name = "RIGHTIE" And s.HasTable Then
            Set pptTable = s.table
            Exit For
        End If
    Next s

    If pptTable Is Nothing Then
        Debug.Print "Table RIGHTIE not found."
        MsgBox "Table RIGHTIE not found on slide.", vbExclamation
        Exit Sub
    End If

    ' Step 3: Adjust the number of rows in LEFTIE
    Dim desiredRows As Long
    desiredRows = Associations_Total - Stronger_Last_Value + 2

    ' Adjust table rows
    Dim existingRows As Long
    existingRows = pptTable.Rows.count

    If existingRows < desiredRows Then
        ' Add missing rows
        Dim addRows As Long
        addRows = desiredRows - existingRows
        Dim j As Long
        For j = 1 To addRows
            pptTable.Rows.Add
        Next j
        Debug.Print "Added " & addRows & " rows to table."
    ElseIf existingRows > desiredRows Then
        ' Remove excess rows
        Dim removeRows As Long
        removeRows = existingRows - desiredRows
        For j = 1 To removeRows
            pptTable.Rows(pptTable.Rows.count).Delete
        Next j
        Debug.Print "Deleted " & removeRows & " rows from table."
    End If

    Debug.Print "Final table row count: " & pptTable.Rows.count
End Sub

