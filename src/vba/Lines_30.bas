Attribute VB_Name = "Lines_30"
Sub Lines_30_Optimized()
    Dim filePath As String
    Dim userName As String
    Dim operatingSystem As String
    Dim fileNum As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim startRow As Long, endRow As Long
    Dim startCol As Long
    Dim values() As String
    Dim strongValuesEnd As String
    Dim weakValuesStart As String

    ' === Get the username from the Environ function ===
    userName = Environ("USER")

    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        ' For Windows
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Define the rows and column to fetch ===
    startRow = 470
    endRow = 471
    startCol = 2 ' Column 2

    ' Open the file
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Read only the relevant rows
    currentRow = 1
    Do Until EOF(fileNum)
        Line Input #fileNum, lineData

        ' Process only rows 470 and 471
        If currentRow = startRow Or currentRow = endRow Then
            values = Split(lineData, ";") ' Split the line by semicolon

            ' Ensure the target column exists
            If UBound(values) >= startCol - 1 Then
                If currentRow = 470 Then
                    strongValuesEnd = Trim(values(startCol - 1)) ' Value at Row 470, Column 2
                ElseIf currentRow = 471 Then
                    weakValuesStart = Trim(values(startCol - 1)) ' Value at Row 471, Column 2
                End If
            End If
        End If

        ' Exit loop early if both rows are processed
        If currentRow >= endRow Then Exit Do

        currentRow = currentRow + 1
    Loop

    ' Close the file
    Close #fileNum

    ' === Print results to Immediate Window ===
    If Len(strongValuesEnd) > 0 Then
        Debug.Print "Row 470, Column 2 (Strong_values_end): " & strongValuesEnd
    Else
        Debug.Print "Row 470, Column 2 (Strong_values_end): [Value not found or empty]"
    End If

    If Len(weakValuesStart) > 0 Then
        Debug.Print "Row 471, Column 2 (Weak_values_start): " & weakValuesStart
    Else
        Debug.Print "Row 471, Column 2 (Weak_values_start): [Value not found or empty]"
    End If
End Sub

