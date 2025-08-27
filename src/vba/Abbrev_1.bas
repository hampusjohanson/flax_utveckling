Attribute VB_Name = "Abbrev_1"

Sub Abbrev_1()
    Dim pptSlide As slide
    Dim userName As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim sourceTable As table
    Dim leftTable As table, rightTable As table
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim tableRow As Integer, tableCol As Integer
    Dim operatingSystem As String
    Dim k As Integer
    Dim copiedText As String ' Declare copiedText once, outside the loop


' === Check Operating System ===
operatingSystem = Application.operatingSystem

' Determine file path based on the operating system
If InStr(operatingSystem, "Macintosh") > 0 Then
    ' For macOS, using ENVIRON function to get the username dynamically
    userName = Environ("USER")
    filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
Else
    ' For Windows
    filePath = "C:\\Local\exported_data_semi.csv"
End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Hard-coded rows and columns ===
    startRow = 392  ' Start at row 392
    endRow = 417    ' End at row 417
    startCol = 1    ' Start at column 1
    endCol = 5      ' End at column 5

    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Create a new table on the current slide
    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable( _
        numRows:=endRow - startRow + 1, _
        NumColumns:=endCol - startCol + 1, _
        left:=50, Top:=50, width:=600, height:=300)

    Set sourceTable = importedTableShape.table

    ' === Read and fill the entire table ===
    rowIndex = 0
    tableRow = 1

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' If the row is within the hard-coded interval
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            tableCol = 1
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    cellValue = Trim(Data(colIndex - 1))
                    sourceTable.cell(tableRow, tableCol).shape.TextFrame.textRange.text = cellValue
                    tableCol = tableCol + 1
                End If
            Next colIndex
            tableRow = tableRow + 1
        End If
    Loop

    Close fileNumber

    ' === Remove borders and fill from each cell ===
    Dim i As Integer, j As Integer
    For i = 1 To sourceTable.Rows.count
        For j = 1 To sourceTable.Columns.count
            With sourceTable.cell(i, j)
                ' Remove borders
                For k = 1 To 4 ' (Top, Left, Bottom, Right)
                    .Borders(k).visible = msoFalse
                Next k
                ' Remove fill color
                .shape.Fill.visible = msoFalse
                ' Standardize text color (black)
                .shape.TextFrame.textRange.Font.color.RGB = RGB(0, 0, 0)
            End With
        Next j
    Next i

    ' === Identify left and right tables on the slide ===
    Dim tableCount As Integer
    tableCount = 0

    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            tableCount = tableCount + 1
            If tableCount = 1 Then
                Set leftTable = shape.table
            ElseIf tableCount = 2 Then
                Set rightTable = shape.table
                Exit For
            End If
        End If
    Next shape

    ' === Copy from new table to left table (columns 1 and 2) ===
    If Not leftTable Is Nothing Then
        Dim sourceRows As Integer, sourceCols As Integer
        sourceRows = sourceTable.Rows.count
        sourceCols = sourceTable.Columns.count

        For rowIndex = 2 To sourceRows
            For colIndex = 1 To 2 ' only columns 1 and 2
                copiedText = sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text
                If rowIndex <= leftTable.Rows.count And colIndex <= leftTable.Columns.count Then
                    leftTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = copiedText
                End If
            Next colIndex
        Next rowIndex
    End If

    ' === Copy from new table to right table (columns 4 and 5) ===
    If Not rightTable Is Nothing Then
        For rowIndex = 2 To sourceTable.Rows.count
            For colIndex = 4 To 5 ' only columns 4 and 5
                copiedText = sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text
                If rowIndex <= rightTable.Rows.count And (colIndex - 3) <= rightTable.Columns.count Then
                    rightTable.cell(rowIndex, colIndex - 3).shape.TextFrame.textRange.text = copiedText
                End If
            Next colIndex
        Next rowIndex
    End If

    ' === Remove rows from left table with "false" or "falskt" in columns 1 or 2 ===
    If Not leftTable Is Nothing Then
        For rowIndex = leftTable.Rows.count To 1 Step -1
            If LCase(Trim(leftTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "*false*" Or _
               LCase(Trim(leftTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "*falskt*" Or _
               LCase(Trim(leftTable.cell(rowIndex, 2).shape.TextFrame.textRange.text)) Like "*false*" Or _
               LCase(Trim(leftTable.cell(rowIndex, 2).shape.TextFrame.textRange.text)) Like "*falskt*" Then
                leftTable.Rows(rowIndex).Delete
            End If
        Next rowIndex
    End If

    ' === Remove rows from right table with "false" or "falskt" in columns 4 or 5 ===
    If Not rightTable Is Nothing Then
        For rowIndex = rightTable.Rows.count To 1 Step -1
            If LCase(Trim(rightTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "*false*" Or _
               LCase(Trim(rightTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "*falskt*" Or _
               LCase(Trim(rightTable.cell(rowIndex, 2).shape.TextFrame.textRange.text)) Like "*false*" Or _
               LCase(Trim(rightTable.cell(rowIndex, 2).shape.TextFrame.textRange.text)) Like "*falskt*" Then
                rightTable.Rows(rowIndex).Delete
            End If
        Next rowIndex
    End If

    ' === Remove borders from empty cells in the new table ===
    Dim tmpRow As Integer, tmpCol As Integer
    For tmpRow = 1 To sourceTable.Rows.count
        For tmpCol = 1 To sourceTable.Columns.count
            If Trim(sourceTable.cell(tmpRow, tmpCol).shape.TextFrame.textRange.text) = "" Then
                With sourceTable.cell(tmpRow, tmpCol)
                    For k = 1 To 4  ' (Top, Left, Bottom, Right)
                        .Borders(k).visible = msoFalse
                    Next k
                End With
            End If
        Next tmpCol
    Next tmpRow

    ' === Delete the newly created table ===
    importedTableShape.Delete
End Sub


