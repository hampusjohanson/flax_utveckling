Attribute VB_Name = "SS_5"
Sub SS_5()
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

    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        userName = Environ("USER")
        If userName = "" Then
            MsgBox "Inget anvŠndarnamn hittades. VŠnligen ange ett anvŠndarnamn.", vbCritical
            Exit Sub
        End If
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        ' For Windows
        filePath = "C:\\Local\\exported_data_semi.csv"
    End If

    ' === Check if the file exists ===
    If Dir(filePath) = "" Then
        MsgBox "CSV-fil inte hittad: " & filePath, vbExclamation
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

    ' === Add "position vs. competitor average" to the imported text in the specified text box ===
    For Each textBoxShape In pptSlide.Shapes
        If textBoxShape.Name = "textruta 5" Then
            If textBoxShape.HasTextFrame Then
                With textBoxShape.TextFrame.textRange
                    .text = cellValue & " position vs. competitor average"
                End With
            End If
            Exit For
        End If
    Next textBoxShape

    ' === Delete the table ===
    tableShape.Delete

End Sub


