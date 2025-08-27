Attribute VB_Name = "Table_Tool_B_1"
Public Leftie_Table_metric_number As Integer
Public Rightie_Table_metric_number As Integer
Public validMetricsLeftRight As Collection  ' Renamed collection to avoid conflict

Sub Table_Tool_B_1()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim operatingSystem As String
    Dim userName As String
    Dim currentRow As Integer
    Dim colIndex As Integer
    Dim metricList As String
    Dim inputChoice As String
    Dim i As Integer
    Dim metricCount As Integer

    Debug.Print "=== Running Table_Tool_B_1 ==="

    ' === Step 1: Initialize validMetricsLeftRight collection ===
    Set validMetricsLeftRight = New Collection

    ' === Step 2: Determine file path (same as Table_Tool_1) ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        userName = Environ("USERNAME")
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' === Step 3: Check if file exists ===
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Step 4: Read CSV and extract valid metrics from row 418 ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    currentRow = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        currentRow = currentRow + 1

        If currentRow = 418 Then
            Data = Split(line, ";")
            metricCount = 0
            For colIndex = 0 To 24
                If UBound(Data) >= colIndex Then
                    Dim value As String
                    value = Trim(Data(colIndex))
                    ' Exclude "False"/"Falskt" (case insensitive)
                    If LCase(value) <> "false" And LCase(value) <> "falskt" And value <> "" Then
                        validMetricsLeftRight.Add value ' Add value to collection
                        metricCount = metricCount + 1
                    End If
                End If
            Next colIndex
            Exit Do
        End If
    Loop
    Close fileNumber

    ' === Step 5: Create numbered list for dropdown ===
    metricList = ""
    ReDim metricArray(1 To metricCount)
    For i = 1 To metricCount
        metricList = metricList & i & ": " & validMetricsLeftRight(i) & vbNewLine
        metricArray(i) = validMetricsLeftRight(i)
    Next i

    ' === Step 6: Ask for Left table's metric selection ===
    inputChoice = InputBox("For similarity calculation: Left table is what?" & vbNewLine & vbNewLine & metricList, "Select Left Table Metric")

    ' === Step 7: If user cancels, exit ===
    If inputChoice = "" Then Exit Sub

    ' === Step 8: Validate input ===
    If Not IsNumeric(inputChoice) Or val(inputChoice) < 1 Or val(inputChoice) > metricCount Then Exit Sub

    ' === Step 9: Store Left table's selection ===
    Leftie_Table_metric_number = val(inputChoice)
    Debug.Print "Left Table Metric Number: " & Leftie_Table_metric_number

    ' === Step 10: Ask for Right table's metric selection ===
    inputChoice = InputBox("For similarity calculation: Right table is what?" & vbNewLine & vbNewLine & metricList, "Select Right Table Metric")

    ' === Step 11: If user cancels, exit ===
    If inputChoice = "" Then Exit Sub

    ' === Step 12: Validate input ===
    If Not IsNumeric(inputChoice) Or val(inputChoice) < 1 Or val(inputChoice) > metricCount Then Exit Sub

    ' === Step 13: Store Right table's selection ===
    Rightie_Table_metric_number = val(inputChoice)
    Debug.Print "Right Table Metric Number: " & Rightie_Table_metric_number

    Debug.Print "Table_Tool_B_1 completed successfully."
End Sub

