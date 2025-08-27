Attribute VB_Name = "AB_1"
Sub AB_1()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim sourceTable As table
    Dim TARGET As table ' Define the variable for the largest table
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim tableRow As Integer
    Dim rightmostTable As shape
    Dim rightMostPosition As Single
    Dim targetRow As Integer
    Dim i As Integer

    ' Dynamically get file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:/Local/exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Adjusted Rows and Columns ===
    startRow = 1  ' Start at row 21
    endRow = 1    ' End at row 40
    startCol = 1   ' Start at column 1
    endCol = 1     ' End at column 6

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

        ' If the row is within the specified interval
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

    ' === Name the table as "SOURCE" by naming the shape that holds the table ===
    importedTableShape.Name = "SOURCE"

    ' === Remove rows from SOURCE that start with "false" or "falskt" ===
    Dim deleteRow As Boolean
    For rowIndex = sourceTable.Rows.count To 1 Step -1 ' Loop through rows in reverse order
        deleteRow = False
        ' Check the value in the first column to determine if it starts with "false" or "falskt"
        If LCase(Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "false*" Or _
           LCase(Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)) Like "falskt*" Then
            deleteRow = True
        End If
        
        If deleteRow Then
            sourceTable.Rows(rowIndex).Delete
        End If
    Next rowIndex

    ' === Remove underscores ("_") from the first column texts in SOURCE ===
    For rowIndex = 1 To sourceTable.Rows.count
        Dim firstColumnText As String
        firstColumnText = sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text
        ' Remove all underscores
        firstColumnText = Replace(firstColumnText, "_", "")
        sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text = firstColumnText
    Next rowIndex

    ' === Find the table on the slide that is furthest to the right (has the largest Left value) ===
    rightMostPosition = -1 ' Start with an impossible left value
    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            If shape.left > rightMostPosition Then
                Set rightmostTable = shape ' Set the rightmost table
                rightMostPosition = shape.left ' Update the position to the new rightmost position
            End If
        End If
    Next shape

    ' === Rename the rightmost table to "TARGET" ===
    If Not rightmostTable Is Nothing Then
        rightmostTable.Name = "TARGET"
        Set TARGET = rightmostTable.table
    End If


    ' Remove the SOURCE table from PowerPoint
    importedTableShape.Delete

End Sub




