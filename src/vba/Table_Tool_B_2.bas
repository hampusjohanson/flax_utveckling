Attribute VB_Name = "Table_Tool_B_2"
Public similarity_score As String ' Global variable to store the similarity score

Sub Table_Tool_B_2()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim operatingSystem As String
    Dim userName As String
    Dim currentRow As Integer
    Dim leftMetric As Integer
    Dim rightMetric As Integer
    Dim column3Value As String

    Debug.Print "=== Running Table_Tool_8 ==="

    ' === Step 1: Set file path based on OS ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        userName = Environ("USERNAME")
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' === Step 2: Check if file exists ===
    If Dir(filePath) = "" Then Exit Sub

    ' === Step 3: Open CSV file and start reading from row 904 ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    currentRow = 0

    ' Skip rows until we reach row 904
    Do While currentRow < 903
        Line Input #fileNumber, line
        currentRow = currentRow + 1
    Loop

    ' === Step 4: Loop through rows starting from row 904 ===
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        currentRow = currentRow + 1
        Data = Split(line, ";")

        ' Check if column 1 matches Leftie_Table_metric_number
        If UBound(Data) >= 2 Then
            If val(Data(0)) = Leftie_Table_metric_number Then
                ' If column 2 matches Rightie_Table_metric_number, get the value in column 3
                If val(Data(1)) = Rightie_Table_metric_number Then
                    column3Value = Trim(Data(2))
                    similarity_score = column3Value ' Save the matching value as similarity_score
                    Debug.Print "Found Value in Column 3: " & similarity_score
                End If
            End If
        End If
    Loop

    Close fileNumber
    Debug.Print "Table_Tool_8 completed successfully."
End Sub

