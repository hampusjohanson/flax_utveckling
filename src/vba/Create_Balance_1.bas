Attribute VB_Name = "Create_Balance_1"
Sub Create_balance_1()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim desktopPath As String
    Dim rowIndex As Integer
    Dim cellValue As Double
    Dim pptSlide As slide
    Dim storeTextBox As shape

    ' Dynamically set file path for macOS or Windows
    If Environ("OS") Like "*Windows*" Then
        desktopPath = "c:/Local/exported_data_semi.csv"
    Else
        desktopPath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If
    filePath = desktopPath

    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Initialize rowIndex
    rowIndex = 0

    ' Read through the file until we reach row 41
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' If we are at row 41
        If rowIndex = 41 Then
            Data = Split(line, ";")
            If UBound(Data) >= 0 Then
                ' Get the value from the first column in row 41
                cellValue = val(Trim(Data(0))) ' The number we want to use
                Debug.Print "Cell value from row 41: " & cellValue
            End If
        End If
    Loop

    ' Close the file
    Close fileNumber

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Check if the text box named "store" already exists, if not, create it
    On Error Resume Next
    Set storeTextBox = pptSlide.Shapes("store")
    On Error GoTo 0

    If storeTextBox Is Nothing Then
        ' Create the text box named "store" if it doesn't exist
        Set storeTextBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50)
        storeTextBox.Name = "store"
    End If

    ' Store the value in the text box
    storeTextBox.TextFrame.textRange.text = cellValue

    ' Debugging: Verify the text box value
    Debug.Print "Stored value in 'store' text box: " & storeTextBox.TextFrame.textRange.text
End Sub

