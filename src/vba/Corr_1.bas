Attribute VB_Name = "Corr_1"
Sub Corr_1()
    Dim pptSlide As slide
    Dim rememberedShape As shape
    Dim importedTableShape As shape
    Dim leftMostTable As table
    Dim selectedTable As table
    Dim filePath As String
    Dim Data() As String
    Dim i As Integer, rowIndex As Integer, colIndex As Integer, foundRow As Integer
    Dim userChoice As String
    Dim fileNumber As Integer
    Dim line As String
    Dim startRow As Integer, endRow As Integer, startCol As Integer, endCol As Integer

    ' === Enable error handling to ignore any issues ===
    On Error Resume Next

    ' Dynamically determine the file path to Desktop based on OS
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:/Local/exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    ' Check if the file exists
    If Dir(filePath) = "" Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    ' === Check if a table is selected ===
    Set rememberedShape = ActiveWindow.Selection.ShapeRange(1)

    If rememberedShape Is Nothing Or Not rememberedShape.HasTable Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    ' === Hardcoded rows and columns ===
    startRow = 162
    endRow = 211
    startCol = 1
    endCol = 6

    ' === Create a new table on the left ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable(numRows:=endRow - startRow + 1, NumColumns:=endCol - startCol + 1, left:=-2000, Top:=50, width:=600, height:=300)

    ' Ensure the table was created successfully
    If importedTableShape Is Nothing Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    Set leftMostTable = importedTableShape.table
    importedTableShape.visible = msoFalse

    ' Read and fill the table from the CSV
    rowIndex = 0
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    With leftMostTable.cell(rowIndex - startRow + 1, colIndex - startCol + 1).shape.TextFrame.textRange
                        .text = Trim(Data(colIndex - 1))
                        .Font.size = 1 ' Set font size to 1
                    End With
                End If
            Next colIndex
        End If
    Loop
    Close fileNumber

    ' Make rows and columns as small as possible
    For i = 1 To leftMostTable.Rows.count
        leftMostTable.Rows(i).height = 0.1
    Next i

    For i = 1 To leftMostTable.Columns.count
        leftMostTable.Columns(i).width = 0.1
    Next i

    ' === Step 3: Process the selected table ===
    Dim dropdownOptions() As String
    ReDim dropdownOptions(1 To leftMostTable.Rows.count)
    For i = 1 To leftMostTable.Rows.count
        dropdownOptions(i) = leftMostTable.cell(i, 1).shape.TextFrame.textRange.text
    Next i

    Load UserForm1
    UserForm1.ComboBox1.Clear
    For i = LBound(dropdownOptions) To UBound(dropdownOptions)
        UserForm1.ComboBox1.AddItem dropdownOptions(i)
    Next i
    UserForm1.Show

    If UserForm1.ComboBox1.ListIndex = -1 Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    userChoice = UserForm1.ComboBox1.value
    Unload UserForm1

    foundRow = -1
    For i = 1 To leftMostTable.Rows.count
        If leftMostTable.cell(i, 1).shape.TextFrame.textRange.text = userChoice Then
            foundRow = i
            Exit For
        End If
    Next i

    If foundRow = -1 Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    rememberedShape.Select
    Set selectedTable = rememberedShape.table

    ' Validate selectedTable before using it
    If selectedTable Is Nothing Then
        Debug.Print "Please select a table first"
        Exit Sub
    End If

    For i = 1 To selectedTable.Columns.count
        selectedTable.cell(1, i).shape.TextFrame.textRange.text = userChoice
    Next i

    For i = 2 To 6
        selectedTable.cell(i, 2).shape.TextFrame.textRange.text = leftMostTable.cell(foundRow, i).shape.TextFrame.textRange.text
    Next i

    importedTableShape.Delete

    ' === Disable error handling after execution ===
    On Error GoTo 0
End Sub

