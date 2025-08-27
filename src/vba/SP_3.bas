Attribute VB_Name = "SP_3"
Public Sub SP_3()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim currentRow As Integer
    Dim Data() As String
    Dim hexColor As String
    Dim r As Integer, g As Integer, b As Integer
    Dim validRowCount As Integer

    ' Get the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart in the slide
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' === Get the username from the Environ function ===
    userName = Environ("USER")
    
    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv" ' Use the username for macOS
    Else
        ' For Windows
        filePath = "C:\Local\exported_data_semi.csv" ' Build the file path for Windows
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the CSV file
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Initialize counters
    currentRow = 1
    validRowCount = 0

    ' Read CSV file (limit to 20 rows)
    Do While Not EOF(fileNumber) And currentRow <= 20
        Line Input #fileNumber, line
        Data = Split(line, ";")

        ' Check if column 6 exists and is valid
        If UBound(Data) >= 5 Then
            hexColor = Trim(Data(5))
            If left(hexColor, 1) = "#" And Len(hexColor) = 7 Then
                validRowCount = validRowCount + 1
                Debug.Print "Valid hex color found: " & hexColor

                ' Apply color directly to chart marker
                If validRowCount <= chartShape.chart.SeriesCollection(1).Points.count Then
                    r = CLng("&H" & Mid(hexColor, 2, 2))
                    g = CLng("&H" & Mid(hexColor, 4, 2))
                    b = CLng("&H" & Mid(hexColor, 6, 2))
                    chartShape.chart.SeriesCollection(1).Points(validRowCount).Format.Fill.ForeColor.RGB = RGB(r, g, b)
                    Debug.Print "Point " & validRowCount & " colored with " & hexColor
                Else
                    Debug.Print "No more points available to color."
                    Exit Do
                End If
            Else
                Debug.Print "Invalid hex color: " & hexColor
            End If
        Else
            Debug.Print "Missing column 6 in row: " & line
        End If

        currentRow = currentRow + 1
    Loop

    Close fileNumber

 
End Sub

