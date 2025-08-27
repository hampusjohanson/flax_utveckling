Attribute VB_Name = "Topp_1"
Sub Topp_1()
    Dim pptSlide As slide
    Dim csvTableShape As shape
    Dim sourceTable As table
    Dim targetShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim tableRowIndex As Integer
    Dim startRow As Integer, endRow As Integer
    Dim selectedColumn As Integer
    Dim operatingSystem As String
    Dim userName As String
    Dim i As Integer
    Dim cellText As String

    ' === Determine the Operating System ===
    operatingSystem = Application.operatingSystem

    ' === Get the username from the environment ===
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' macOS
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        ' Windows
        userName = Environ("USERNAME")
        filePath = "C:\\Local\\exported_data_semi.csv"
    End If

    ' Validate username
    If Trim(userName) = "" Then
        MsgBox "No username was set. Please check your system and try again.", vbCritical
        Exit Sub
    End If

    ' Validate File Path
    If Dir(filePath) = "" Then
        MsgBox "The file was not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Ensure selectedTopDriver is Valid ===
    Select Case selectedTopDriver
        Case "Sales premium"
            selectedColumn = 1
        Case "Volume premium"
            selectedColumn = 2
        Case "Price premium"
            selectedColumn = 3
        Case Else
            MsgBox "Invalid selection for top driver: [" & selectedTopDriver & "]. Please select one of: Sales premium, Volume premium, or Price premium.", vbCritical
            Exit Sub
    End Select

    ' === Open CSV File ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' === Set Row Range ===
    startRow = 418
    endRow = 468

    ' === Create Table with One Column at Specified Location ===
    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable( _
        numRows:=endRow - startRow + 1, _
        NumColumns:=1, _
        left:=28.35, Top:=582.24, width:=600, height:=300)
    importedTableShape.Name = "SOURCE"
    Set sourceTable = importedTableShape.table

    ' === Read and Fill the Table ===
    rowIndex = 0
    tableRowIndex = 1 ' Table rows start at 1
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' Import only the selected column within the specified row range
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            On Error Resume Next
            sourceTable.cell(tableRowIndex, 1).shape.TextFrame.textRange.text = Trim(Data(selectedColumn - 1))
            On Error GoTo 0
            tableRowIndex = tableRowIndex + 1 ' Move to the next table row
        End If
    Loop

    ' Close the file
    Close fileNumber

    ' === Remove "FALSE" or "FALSKT" Rows from SOURCE ===
    For i = sourceTable.Rows.count To 1 Step -1
        On Error Resume Next
        cellText = Trim(LCase(sourceTable.cell(i, 1).shape.TextFrame.textRange.text))
        If cellText Like "false*" Or cellText Like "falskt*" Then
            sourceTable.Rows(i).Delete ' Delete the entire row
        End If
        On Error GoTo 0
    Next i

    ' === Remove the First Row of SOURCE ===
    On Error Resume Next
    sourceTable.Rows(1).Delete
    On Error GoTo 0

    ' === Rename the Existing Table to TARGET ===
    On Error Resume Next
    Set targetShape = pptSlide.Shapes("TARGET")
    If targetShape Is Nothing Then
        For Each shape In pptSlide.Shapes
            If shape.HasTable Then
                shape.Name = "TARGET"
                Set targetShape = shape
                Exit For
            End If
        Next shape
    End If
    On Error GoTo 0

    If targetShape Is Nothing Then
        MsgBox "No table found to rename as TARGET.", vbCritical
        Exit Sub
    End If

    Set targetTable = targetShape.table

    ' === Paste Rows from SOURCE to TARGET ===
    For i = 1 To 50
        If i > sourceTable.Rows.count Then Exit For
        Select Case True
            Case i <= 10
                targetTable.cell(i + 1, 2).shape.TextFrame.textRange.text = sourceTable.cell(i, 1).shape.TextFrame.textRange.text
            Case i <= 20
                targetTable.cell(i - 10 + 1, 5).shape.TextFrame.textRange.text = sourceTable.cell(i, 1).shape.TextFrame.textRange.text
            Case i <= 30
                targetTable.cell(i - 20 + 1, 8).shape.TextFrame.textRange.text = sourceTable.cell(i, 1).shape.TextFrame.textRange.text
            Case Else
                targetTable.cell(i - 30 + 1, 11).shape.TextFrame.textRange.text = sourceTable.cell(i, 1).shape.TextFrame.textRange.text
        End Select
    Next i

    ' === Delete SOURCE Table ===
    importedTableShape.Delete

    ' === Erase Cells in TARGET Columns 10 and 11 ===
    For i = 1 To targetTable.Rows.count
        On Error Resume Next
        cellText = Trim(LCase(targetTable.cell(i, 11).shape.TextFrame.textRange.text))
        If cellText = "" Or cellText Like "false*" Or cellText Like "falskt*" Then
            targetTable.cell(i, 11).shape.TextFrame.textRange.text = "" ' Erase text in column 11
            targetTable.cell(i, 10).shape.TextFrame.textRange.text = "" ' Erase corresponding text in column 10
        End If
        On Error GoTo 0
    Next i

    ' === Add New Logic for Updating Row 1 Texts ===
    Select Case selectedTopDriver
        Case "Sales premium"
            targetTable.cell(1, 1).shape.TextFrame.textRange.text = "Highest impact on sales premium"
            targetTable.cell(1, 4).shape.TextFrame.textRange.text = "Higher impact on sales premium"
            targetTable.cell(1, 7).shape.TextFrame.textRange.text = "Weaker impact on sales premium"
            targetTable.cell(1, 10).shape.TextFrame.textRange.text = "Weakest impact on sales premium"
        Case "Volume premium"
            targetTable.cell(1, 1).shape.TextFrame.textRange.text = "Highest impact on volume premium"
            targetTable.cell(1, 4).shape.TextFrame.textRange.text = "Higher impact on volume premium"
            targetTable.cell(1, 7).shape.TextFrame.textRange.text = "Weaker impact on volume premium"
            targetTable.cell(1, 10).shape.TextFrame.textRange.text = "Weakest impact on volume premium"
        Case "Price premium"
            targetTable.cell(1, 1).shape.TextFrame.textRange.text = "Highest impact on price premium"
            targetTable.cell(1, 4).shape.TextFrame.textRange.text = "Higher impact on price premium"
            targetTable.cell(1, 7).shape.TextFrame.textRange.text = "Weaker impact on price premium"
            targetTable.cell(1, 10).shape.TextFrame.textRange.text = "Weakest impact on price premium"
        Case Else
            MsgBox "Unknown driver: " & selectedTopDriver, vbExclamation
    End Select
End Sub




