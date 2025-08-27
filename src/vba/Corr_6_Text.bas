Attribute VB_Name = "Corr_6_Text"
Sub Corr_6_Text()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim operatingSystem As String
    Dim userName As String
    Dim brandName As String
    Dim currentRow As Integer
    Dim textBox As shape
    Dim foundTextBox As Boolean

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide set."

    ' === Check OS and determine file path ===
    operatingSystem = Application.operatingSystem
    Debug.Print "Operating system: " & operatingSystem

    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        userName = Environ("USERNAME")
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    Debug.Print "File path set: " & filePath

    ' Check if file exists
    If Dir(filePath) = "" Then
        Debug.Print "File not found: " & filePath
        Exit Sub
    End If

    ' === Open the CSV file ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    Debug.Print "CSV file opened."

    ' Loop to find row 850
    currentRow = 0
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        currentRow = currentRow + 1

        ' Extract brand name from row 850, column 1
        If currentRow = 850 Then
            Data = Split(line, ";")
            If UBound(Data) >= 0 Then
                brandName = Trim(Data(0))
            Else
                brandName = "Unknown"
            End If
            Debug.Print "Brand name from Row 850: " & brandName
            Exit Do
        End If
    Loop
    Close fileNumber
    Debug.Print "CSV file closed."

    ' === Find "Rubrik 2" and update text ===
    foundTextBox = False
    For Each textBox In pptSlide.Shapes
        If textBox.HasTextFrame Then
            If textBox.Name = "Rubrik 2" Then
                textBox.TextFrame.textRange.text = "Correlations to the strongest associations of " & brandName
                Debug.Print "Updated 'Rubrik 2' with: " & brandName
                foundTextBox = True
                Exit For
            End If
        End If
    Next textBox

    If Not foundTextBox Then Debug.Print "Rubrik 2 not found."
    
    Debug.Print "UpdateRubrik2WithBrand completed."
End Sub

