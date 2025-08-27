Attribute VB_Name = "Lines_0"
Sub Lines_0()
    Dim originalFilePath As String
    Dim cleanedFilePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim rowIndex As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim newFileNumber As Integer
    Dim isMac As Boolean

    ' Determine the operating system
    isMac = Application.operatingSystem Like "*Macintosh*"

    ' Define file paths based on the OS
    If isMac Then
        originalFilePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
        cleanedFilePath = "/Users/" & Environ("USER") & "/Desktop/line_chart_data_csv.csv"
    Else
        originalFilePath = "C:\Local\exported_data_semi.csv"
        cleanedFilePath = "C:\Local\line_chart_data_csv.csv"
    End If

    ' Check if the original file exists
    If Dir(originalFilePath) = "" Then
        MsgBox "Original file not found: " & originalFilePath, vbExclamation
        Exit Sub
    End If

    ' Open the original file to read data
    fileNumber = FreeFile
    Open originalFilePath For Input As fileNumber

    ' Create a new file to write the cleaned data
    newFileNumber = FreeFile
    Open cleanedFilePath For Output As newFileNumber

    ' Set start and end rows for the data we want to keep
    startRow = 735
    endRow = 785
    startCol = 1
    endCol = 21

    rowIndex = 1

    ' Loop through the original CSV file
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        Data = Split(line, ";")

        ' Check if the row is within the desired range
        If rowIndex >= startRow And rowIndex <= endRow Then
            ' Write only the first 21 columns for this row
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    ' Write the column value to the new cleaned file
                    Print #newFileNumber, Data(colIndex - 1) & ";";
                End If
            Next colIndex
            Print #newFileNumber, "" ' Newline after each row
        End If

        rowIndex = rowIndex + 1
    Loop

    ' Close the files
    Close fileNumber
    Close newFileNumber

End Sub

