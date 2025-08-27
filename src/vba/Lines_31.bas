Attribute VB_Name = "Lines_31"
Sub Lines_31()
    ' SetVerticalAxisWithMapping- Stronger
    Dim chartShape As shape
    Dim chartObject As chart
    Dim maxValue As Double
    Dim minValue As Double
    Dim Stronger_Last_Value As Double
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

    ' Open the CSV file and read Stronger_Last_Value
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        
        ' Look for row 470 (Stronger_Last_Value)
        If currentRow = 470 Then
            dataArray = Split(lineData, ";")
            Stronger_Last_Value = CDbl(dataArray(1))
            Exit Do
        End If
    Loop
    Close #fileNumber

    ' Validate Stronger_Last_Value
    If Stronger_Last_Value < 1 Or Stronger_Last_Value > 50 Then
        MsgBox "Invalid Stronger_Last_Value: " & Stronger_Last_Value, vbExclamation
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
    minValue = 51 - 1 ' Map 1 to 50
    maxValue = 51 - Stronger_Last_Value ' Map Stronger_Last_Value accordingly

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

