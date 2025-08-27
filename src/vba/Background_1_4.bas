Attribute VB_Name = "Background_1_4"
Sub Background_1_4()
    ' Step 1: Read Stronger_Last_Value from row 470, as before.
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim dataArray() As String
    Dim Stronger_Last_Value As Double

    ' Dynamically get the file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:\\Local\\exported_data_semi.csv"
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

    ' Open the CSV file and read Stronger_Last_Value (row 470)
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
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

    Debug.Print "Final Stronger_Last_Value: " & Stronger_Last_Value

    ' Step 2: Reference the active slide and find the table named LEFTIE
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide reference set."
    
    Dim pptTable As table
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

    ' Step 3: Open the file again to read rows 573 to (573 + Stronger_Last_Value - 1)
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0

    ' Skip rows up to 572
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        If currentRow = 572 Then Exit Do
    Loop

    ' Next line read is row 573; read Stronger_Last_Value lines.
    ' We'll extract CSV columns 1, 2, 3, and 4 => dataArray(0), dataArray(1), dataArray(2), dataArray(3).

    Dim linesNeeded As Long
    linesNeeded = Stronger_Last_Value

    Dim i As Long
    For i = 1 To linesNeeded
        If EOF(fileNumber) Then
            Debug.Print "Reached end of file before reading all needed rows." & _
                        " i=" & i & ", linesNeeded=" & linesNeeded
            Exit For
        End If
        
        Line Input #fileNumber, lineData
        dataArray = Split(lineData, ";")

        Dim col1 As String, col2 As String, col3 As String, col4 As String
        col1 = ""
        col2 = ""
        col3 = ""
        col4 = ""

        If UBound(dataArray) >= 0 Then col1 = dataArray(0)
        If UBound(dataArray) >= 1 Then col2 = dataArray(1)
        If UBound(dataArray) >= 2 Then col3 = dataArray(2)
        If UBound(dataArray) >= 3 Then col4 = dataArray(3)

        ' Step 4: Paste into the table in row (3 + (i - 1)), columns 1, 2, 3, 4
        Dim tableRow As Long
        tableRow = 3 + (i - 1)

        ' Safety check: ensure the table has enough rows/columns
        If tableRow <= pptTable.Rows.count And pptTable.Columns.count >= 4 Then
            pptTable.cell(tableRow, 1).shape.TextFrame.textRange.text = col1
            pptTable.cell(tableRow, 2).shape.TextFrame.textRange.text = col2
            pptTable.cell(tableRow, 3).shape.TextFrame.textRange.text = col3
            pptTable.cell(tableRow, 4).shape.TextFrame.textRange.text = col4

            Debug.Print "Pasted data into Table(row=" & tableRow & ", cols=1-4) => (" _
                        & col1 & ", " & col2 & ", " & col3 & ", " & col4 & ")"
        Else
            Debug.Print "Table does not have enough rows/columns for row=" & tableRow
        End If
    Next i

    Close #fileNumber
End Sub


