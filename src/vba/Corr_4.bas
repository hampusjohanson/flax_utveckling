Attribute VB_Name = "Corr_4"
Sub Corr_4()
    Dim pptSlide As slide
    Dim shapeItem As shape
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim operatingSystem As String
    Dim userName As String
    Dim tableShape As shape
    Dim sourceTable As table
    Dim foundTables As New Collection
    Dim TableIndex As Integer
    Dim currentRow As Integer
    Dim rowData(1 To 6, 1 To 6) As String ' Store CSV values

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide set."

    ' === Find and rename tables ===
    For Each shapeItem In pptSlide.Shapes
        If shapeItem.HasTable Then
            shapeItem.Name = "Table" & foundTables.count + 1
            foundTables.Add shapeItem
            Debug.Print "Renamed and added table: " & shapeItem.Name
        End If
    Next shapeItem

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

    ' Loop to read data from rows
    currentRow = 0
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        currentRow = currentRow + 1

        If currentRow >= 640 And currentRow <= 645 Then
            Data = Split(line, ";")
            Debug.Print "Reading row " & currentRow & ": " & line

            ' Store values in rowData
            Dim colIndex As Integer
            For colIndex = 1 To 6
                If UBound(Data) >= colIndex - 1 Then
                    rowData(currentRow - 639, colIndex) = Trim(Data(colIndex - 1))
                Else
                    rowData(currentRow - 639, colIndex) = ""
                End If
            Next colIndex
        End If
    Loop
    Close fileNumber
    Debug.Print "CSV file closed."

    ' === Insert data into tables ===
    On Error Resume Next
    For TableIndex = 1 To foundTables.count
        Set tableShape = foundTables(TableIndex)
        Set sourceTable = tableShape.table

        ' Insert first column (header) into first row, first column
        sourceTable.cell(1, 1).shape.TextFrame.textRange.text = rowData(TableIndex, 1)
        Debug.Print "Table " & TableIndex & ", Row 1, Col 1: " & rowData(TableIndex, 1)

        ' Insert remaining columns into correct rows and columns
        Dim r As Integer
        For r = 1 To 6
            sourceTable.cell(r + 1, 2).shape.TextFrame.textRange.text = rowData(TableIndex, r + 1)
            Debug.Print "Table " & TableIndex & ", Row " & r + 1 & ", Col 2: " & rowData(TableIndex, r + 1)
        Next r
    Next TableIndex
    On Error GoTo 0

End Sub



