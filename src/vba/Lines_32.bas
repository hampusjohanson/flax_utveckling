Attribute VB_Name = "Lines_32"
Sub Lines_32()
    ' SetVerticalAxisWithMapping - WEAKER
    Dim chartShape As shape
    Dim chartObject As chart
    Dim maxValue As Double
    Dim minValue As Double
    Dim Weaker_First_Value As Double
    Dim Last_Value As Double
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim dataArray() As String

    ' Dynamically get the file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:\Local\exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the CSV file and read Weaker_First_Value and Last_Value
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        
        ' Look for row 471 (Weaker_First_Value)
        If currentRow = 471 Then
            dataArray = Split(lineData, ";")
            Weaker_First_Value = CDbl(dataArray(1))
        End If
        
        ' Look for row 469 (Last_Value)
        If currentRow = 469 Then
            dataArray = Split(lineData, ";")
            Last_Value = CDbl(dataArray(1))
        End If
    Loop
    Close #fileNumber

    ' Validate Weaker_First_Value and Last_Value
    If Weaker_First_Value < 1 Or Weaker_First_Value > 50 Then
        MsgBox "Invalid Weaker_First_Value: " & Weaker_First_Value, vbExclamation
        Exit Sub
    End If
    If Last_Value < 1 Or Last_Value > 50 Then
        MsgBox "Invalid Last_Value: " & Last_Value, vbExclamation
        Exit Sub
    End If

    ' Check if a shape is selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera diagram tack.", vbExclamation
        Exit Sub
    End If

    ' Check if the selected shape is a chart
    Set chartShape = ActiveWindow.Selection.ShapeRange(1)
    If Not chartShape.hasChart Then
        MsgBox "Markera ett giltigt diagram tack.", vbExclamation
        Exit Sub
    End If

    ' Get the chart object
    Set chartObject = chartShape.chart

    ' Default values
    minValue = 51 - Weaker_First_Value ' Map Weaker_First_Value
    maxValue = 51 - Last_Value         ' Map Last_Value

    ' Ensure max > min (swap if needed)
    If maxValue < minValue Then
        Dim temp As Double
        temp = maxValue
        maxValue = minValue
        minValue = temp
    End If

    ' Set the Min and Max values for the vertical axis
    With chartObject.Axes(xlValue)
        .MinimumScale = minValue
        .MaximumScale = maxValue
    End With

    End Sub

