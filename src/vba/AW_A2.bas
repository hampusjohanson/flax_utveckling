Attribute VB_Name = "AW_A2"
Sub AW_A2()
    On Error Resume Next ' Ignorera alla fel

    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim tableShape As shape
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim operatingSystem As String
    Dim userName As String
    Dim i As Long
    Dim allLines() As String
    Dim lineCount As Long

    userName = Environ("USER")
    operatingSystem = Application.operatingSystem

    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    If Dir(filePath) = "" Then Exit Sub

    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        lineCount = lineCount + 1
    Loop
    Close fileNumber

    ReDim allLines(1 To lineCount)

    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    For i = 1 To lineCount
        Line Input #fileNumber, allLines(i)
    Next i
    Close fileNumber

    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = pptSlide.Shapes("SOURCE")
    If tableShape Is Nothing Then Exit Sub
    If tableShape.HasTable = msoFalse Then Exit Sub

    rowIndex = 1
    For i = 664 To 684
        If i > lineCount Then Exit For

        line = allLines(i)
        If Trim(line) <> "" Then
            Data = Split(line, ";")
            For colIndex = 1 To 6
                If colIndex <= UBound(Data) + 1 Then
                    cellValue = Trim(Data(colIndex - 1))
                    If right(cellValue, 1) = "_" Or right(cellValue, 1) = "?" Then
                        cellValue = left(cellValue, Len(cellValue) - 1)
                    End If
                    tableShape.table.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = cellValue
                End If
            Next colIndex
        End If
        rowIndex = rowIndex + 1
    Next i
End Sub

