Attribute VB_Name = "SP_02"
Sub SP_02()
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim sourceTable As table
    Dim TARGET As table ' Define the variable for the largest table
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim tableRow As Integer
    Dim targetRow As Integer

    ' Dynamically get the file path to the desktop based on the OS
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
    startRow = 1   ' Start at row 1
    endRow = 20    ' End at row 20
    startCol = 1   ' Start at column 1
    endCol = 4     ' End at column 4

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the table renamed as "TARGET" on the slide
    On Error Resume Next
    Set TARGET = pptSlide.Shapes("TARGET").table
    On Error GoTo 0

    ' Check if TARGET table is found
    If TARGET Is Nothing Then
        MsgBox "No table named 'TARGET' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Initialize the row number for pasting data into TARGET
    targetRow = 2 ' Start from row 2 in TARGET (since row 1 will be used for the header)

    ' === Read and paste data directly into TARGET table ===
    rowIndex = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' If the row is within the specified interval
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")

            ' Check if the row contains "false" or "falskt" in the first column and skip it
            If LCase(Trim(Data(0))) Like "false*" Or LCase(Trim(Data(0))) Like "falskt*" Then
                GoTo SkipRow
            End If

            ' Dynamically add rows if needed
            If targetRow > TARGET.Rows.count Then
                Do While targetRow > TARGET.Rows.count
                    TARGET.Rows.Add
                Loop
            End If

            ' Paste the data into the TARGET table
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    cellValue = Trim(Data(colIndex - 1))
                    TARGET.cell(targetRow, colIndex - startCol + 1).shape.TextFrame.textRange.text = cellValue
                End If
            Next colIndex
            targetRow = targetRow + 1
        End If
SkipRow:
    Loop

    Close fileNumber

    ' === Remove underscores ("_") from the first column texts in TARGET ===
    For rowIndex = 1 To TARGET.Rows.count
        Dim firstColumnText As String
        firstColumnText = TARGET.cell(rowIndex, 1).shape.TextFrame.textRange.text
        ' Remove all underscores
        firstColumnText = Replace(firstColumnText, "_", "")
        TARGET.cell(rowIndex, 1).shape.TextFrame.textRange.text = firstColumnText
    Next rowIndex

    ' === Remove empty rows in TARGET ===
    For rowIndex = TARGET.Rows.count To 1 Step -1 ' Loop through rows in reverse order
        If IsRowEmpty(TARGET, rowIndex) Then
            TARGET.Rows(rowIndex).Delete
        End If
    Next rowIndex

    ' === Delete extra rows from the bottom of the table if necessary ===
    If targetRow <= TARGET.Rows.count Then
        For tableRow = TARGET.Rows.count To targetRow Step -1
            TARGET.Rows(tableRow).Delete
        Next tableRow
    End If

End Sub

' Function to check if a row is empty
Function IsRowEmpty(tbl As table, row As Integer) As Boolean
    Dim colIndex As Integer
    IsRowEmpty = True ' Assume the row is empty initially

    For colIndex = 1 To tbl.Columns.count
        If Trim(tbl.cell(row, colIndex).shape.TextFrame.textRange.text) <> "" Then
            IsRowEmpty = False ' If any cell in the row is not empty, it's not an empty row
            Exit Function
        End If
    Next colIndex
End Function

