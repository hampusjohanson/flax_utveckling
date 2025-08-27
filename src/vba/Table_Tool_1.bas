Attribute VB_Name = "Table_Tool_1"
Public metric_decision As String
Public row_decision As Integer
Public number_metric As Integer ' Global variable for the metric number
Public validMetrics As Collection ' Global collection for valid metrics

Sub Table_Tool_1()
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
    Dim shapeItem As shape

    ' === Step 1: Check if a table shape is selected ===
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shapeItem = ActiveWindow.Selection.ShapeRange(1)
        
        ' Check if the shape contains a table
        If shapeItem.HasTable Then
            Debug.Print "Table selected: " & shapeItem.Name
        Else
            MsgBox "Please select a valid table first.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Please select a valid table first.", vbExclamation
        Exit Sub
    End If

    ' === Step 2: Determine file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        userName = Environ("USERNAME")
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' === Step 3: Check if file exists ===
    If Dir(filePath) = "" Then Exit Sub

    ' === Step 4: Initialize Collection ===
    Set validMetrics = New Collection

    ' === Step 5: Read CSV and extract valid metrics from row 418 ===
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
                        validMetrics.Add value ' Add value to collection
                        metricCount = metricCount + 1
                    End If
                End If
            Next colIndex
            Exit Do
        End If
    Loop
    Close fileNumber

    ' === Step 6: Create numbered list for dropdown ===
    metricList = ""
    ReDim metricArray(1 To metricCount)
    For i = 1 To metricCount
        metricList = metricList & i & ": " & validMetrics(i) & vbNewLine
        metricArray(i) = validMetrics(i)
    Next i

    ' === Step 7: Ask user for metric selection ===
    inputChoice = InputBox("Top values of what?" & vbNewLine & vbNewLine & metricList, "Select Metric")

    ' === Step 8: If user cancels, exit ===
    If inputChoice = "" Then Exit Sub

    ' === Step 9: Validate input ===
    If Not IsNumeric(inputChoice) Or val(inputChoice) < 1 Or val(inputChoice) > metricCount Then Exit Sub

    ' === Step 10: Store valid selection ===
    metric_decision = metricArray(val(inputChoice))
    number_metric = val(inputChoice) ' Store the selected number as "number_metric"
    Debug.Print "Metric Decision: " & metric_decision
    Debug.Print "Metric Number: " & number_metric ' Debugging the stored number

    ' === Step 11: Ask user for row count ===
    inputChoice = InputBox("Insert how many rows? (1-50)", "Row Count", 10)

    ' === Step 12: If user cancels, exit ===
    If inputChoice = "" Then Exit Sub

    ' === Step 13: Convert and validate numeric input ===
    row_decision = val(inputChoice)
    If Not IsNumeric(row_decision) Or row_decision < 1 Or row_decision > 50 Then Exit Sub

    ' === Step 14: Store row count ===
    Debug.Print "Row Decision: " & row_decision
End Sub

