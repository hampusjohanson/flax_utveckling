Attribute VB_Name = "Background_1_6"
Sub Background_1_6()
    ' Step 1: Read Stronger_Last_Value and Associations_Total from the CSV
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

    ' Step 2: Reference the slide and find the table named RIGHTIE
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

    ' Step 3: Open the file again to read and paste data from row (573 + Stronger_Last_Value) to (573 + Associations_Total - 1)
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0
    
    Dim startRow As Long, endRow As Long
    startRow = 573 + Stronger_Last_Value
    endRow = 573 + Associations_Total - 1
    
    ' Skip rows up to startRow - 1
    Do While Not EOF(fileNumber) And currentRow < startRow - 1
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
    Loop
    
    Dim tableRow As Long
    tableRow = 3 ' Start pasting from row 3 in RIGHTIE
    
    Do While Not EOF(fileNumber) And currentRow <= endRow
        Line Input #fileNumber, lineData
        dataArray = Split(lineData, ";")
        
        If tableRow <= pptTable.Rows.count And pptTable.Columns.count >= 4 Then ' Ensure the row exists
            If UBound(dataArray) >= 3 Then ' Ensure enough columns exist
                pptTable.cell(tableRow, 1).shape.TextFrame.textRange.text = dataArray(0) ' Column 1
                pptTable.cell(tableRow, 2).shape.TextFrame.textRange.text = dataArray(1) ' Column 2
                pptTable.cell(tableRow, 3).shape.TextFrame.textRange.text = dataArray(2) ' Column 3
                pptTable.cell(tableRow, 4).shape.TextFrame.textRange.text = dataArray(3) ' Column 4
                Debug.Print "Pasted into RIGHTIE(row=" & tableRow & ", cols=1-4): " & dataArray(0) & ", " & dataArray(1) & ", " & dataArray(2) & ", " & dataArray(3)
            End If
        Else
            Debug.Print "Skipping row: " & tableRow & " as it exceeds table size."
        End If
        
        tableRow = tableRow + 1
        currentRow = currentRow + 1
    Loop
    
    Close #fileNumber
End Sub

