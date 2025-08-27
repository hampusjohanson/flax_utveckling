Attribute VB_Name = "Sub_9"
Sub Sub_9()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim rowIndex As Long
    Dim cellValue As String
    Dim operatingSystem As String
    Dim tableShape As shape
    Dim textBoxShape As shape
    Dim userName As String
    Dim shapeFound As Boolean

    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        userName = Environ("USER")
        If userName = "" Then
            MsgBox "No username found. Please provide a username.", vbCritical
            Exit Sub
        End If
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        ' For Windows
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' === Check if the file exists ===
    If Dir(filePath) = "" Then
        MsgBox "CSV file not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Read data from CSV ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    rowIndex = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        If rowIndex = 341 Then
            Data = Split(line, ";")
            If UBound(Data) >= 0 Then
                cellValue = Replace(Trim(Data(0)), "_", "") ' Remove "_"
                cellValue = Replace(cellValue, "?", "") ' Remove all "?"
            Else
                cellValue = ""
            End If
            Exit Do
        End If
    Loop

    Close fileNumber

    ' === Insert table on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = pptSlide.Shapes.AddTable(1, 1, 100, 100, 100, 50) ' Position and size of the table

    ' Insert the value into the single cell
    With tableShape.table
        .cell(1, 1).shape.TextFrame.textRange.text = cellValue
    End With

    ' === Find "textruta 5" or a textbox at the given coordinates ===
    shapeFound = False
    For Each textBoxShape In pptSlide.Shapes
        If textBoxShape.HasTextFrame Then
            ' Check if it's explicitly "textruta 5"
            If textBoxShape.Name = "textruta 5" Then
                textBoxShape.Name = "brand_x" ' Rename the textbox
                shapeFound = True
            End If
            
            ' Check if the shape is at the expected position (4.27 cm, 16.13 cm)
            If Not shapeFound And Abs(textBoxShape.left - 120) < 5 And Abs(textBoxShape.Top - 457) < 5 Then
                ' Approximate position check in points (~1 cm ˜ 28.35 points)
                textBoxShape.Name = "brand_x"
                shapeFound = True
            End If
            
            ' If we found the correct shape, update its text
            If shapeFound Then
                textBoxShape.TextFrame.textRange.text = cellValue & " position vs. competitor average"
                Exit For
            End If
        End If
    Next textBoxShape

    If Not shapeFound Then
        MsgBox "Textbox not found at expected position.", vbExclamation
    Else
        Debug.Print "Textbox updated successfully."
    End If

    ' === Delete the temporary table ===
    tableShape.Delete

End Sub

