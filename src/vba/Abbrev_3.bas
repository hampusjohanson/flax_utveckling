Attribute VB_Name = "Abbrev_3"
Sub Abbrev_11()
    Dim pptSlide As slide
    Dim userName As String, filePath As String, operatingSystem As String
    Dim fileNumber As Integer, line As String, Data() As String
    Dim importedTableShape As shape, sourceTable As table
    Dim rowIndex As Integer
    Dim startRow As Integer: startRow = 392
    Dim endRow As Integer: endRow = 417
    Dim totalRows As Integer: totalRows = endRow - startRow + 1

    ' === File path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Create temp table with 5 columns ===
    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable( _
        numRows:=totalRows, NumColumns:=5, _
        left:=50, Top:=50, width:=600, height:=300)
    importedTableShape.Name = "ImportedTable"
    Set sourceTable = importedTableShape.table

    ' === Read CSV and fill table ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    rowIndex = 0
    Dim tableRow As Integer: tableRow = 1

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            If UBound(Data) >= 4 Then
                sourceTable.cell(tableRow, 1).shape.TextFrame.textRange.text = Trim(Data(0)) ' Abbr left
                sourceTable.cell(tableRow, 2).shape.TextFrame.textRange.text = Trim(Data(1)) ' Full left
                sourceTable.cell(tableRow, 3).shape.TextFrame.textRange.text = Trim(Data(2)) ' False/True
                sourceTable.cell(tableRow, 4).shape.TextFrame.textRange.text = Trim(Data(3)) ' Abbr right
                sourceTable.cell(tableRow, 5).shape.TextFrame.textRange.text = Trim(Data(4)) ' Full right
                tableRow = tableRow + 1
            End If
        End If
    Loop
    Close fileNumber
End Sub

Sub Abbrev_12()
    Dim sourceTable As table
    Set sourceTable = ActiveWindow.View.slide.Shapes("ImportedTable").table
    Dim i As Integer, j As Integer, k As Integer

    For i = 1 To sourceTable.Rows.count
        For j = 1 To sourceTable.Columns.count
            With sourceTable.cell(i, j)
                For k = 1 To 4: .Borders(k).visible = msoFalse: Next k
                .shape.Fill.visible = msoFalse
                .shape.TextFrame.textRange.Font.color.RGB = RGB(0, 0, 0)
                If Trim(.shape.TextFrame.textRange.text) = "" Then
                    For k = 1 To 4: .Borders(k).visible = msoFalse: Next k
                End If
            End With
        Next j
    Next i
End Sub


Sub Abbrev_13()
    Dim shape As shape, count As Integer
    Dim leftTable As table, rightTable As table
    Dim slide As slide: Set slide = ActiveWindow.View.slide

    For Each shape In slide.Shapes
        If shape.HasTable Then
            count = count + 1
            If count = 1 Then
                shape.Name = "LeftTableShape"
                slide.Tags.Add "LeftTableName", shape.Name
            ElseIf count = 2 Then
                shape.Name = "RightTableShape"
                slide.Tags.Add "RightTableName", shape.Name
                Exit For
            End If
        End If
    Next shape
End Sub




Sub Abbrev_14()
    Dim sourceTable As table
    Dim leftTable As table, rightTable As table
    Dim leftShapeName As String, rightShapeName As String
    Dim slide As slide: Set slide = ActiveWindow.View.slide

    Set sourceTable = slide.Shapes("ImportedTable").table
    leftShapeName = slide.Tags("LeftTableName")
    rightShapeName = slide.Tags("RightTableName")

    Set leftTable = slide.Shapes(leftShapeName).table
    Set rightTable = slide.Shapes(rightShapeName).table

    Dim neededRows As Integer
    neededRows = sourceTable.Rows.count

    Do While leftTable.Rows.count < neededRows
        leftTable.Rows.Add
    Loop
    Do While rightTable.Rows.count < neededRows
        rightTable.Rows.Add
    Loop
End Sub

Sub Abbrev_15()
    Dim sourceTable As table, leftTable As table, rightTable As table
    Dim slide As slide: Set slide = ActiveWindow.View.slide
    Set sourceTable = slide.Shapes("ImportedTable").table
    Set leftTable = slide.Shapes(slide.Tags("LeftTableName")).table
    Set rightTable = slide.Shapes(slide.Tags("RightTableName")).table

    Dim r As Integer
    For r = 1 To sourceTable.Rows.count
        leftTable.cell(r, 1).shape.TextFrame.textRange.text = sourceTable.cell(r, 1).shape.TextFrame.textRange.text
        leftTable.cell(r, 2).shape.TextFrame.textRange.text = sourceTable.cell(r, 2).shape.TextFrame.textRange.text
        rightTable.cell(r, 1).shape.TextFrame.textRange.text = sourceTable.cell(r, 4).shape.TextFrame.textRange.text
        rightTable.cell(r, 2).shape.TextFrame.textRange.text = sourceTable.cell(r, 5).shape.TextFrame.textRange.text
    Next r
End Sub

Sub Abbrev_16()
    Dim slide As slide: Set slide = ActiveWindow.View.slide
    Dim leftTable As table, rightTable As table
    Set leftTable = slide.Shapes(slide.Tags("LeftTableName")).table
    Set rightTable = slide.Shapes(slide.Tags("RightTableName")).table

    Dim i As Integer

    For i = leftTable.Rows.count To 1 Step -1
        If LCase(leftTable.cell(i, 1).shape.TextFrame.textRange.text) Like "*false*" Or _
           LCase(leftTable.cell(i, 2).shape.TextFrame.textRange.text) Like "*false*" Or _
           LCase(leftTable.cell(i, 1).shape.TextFrame.textRange.text) Like "*falskt*" Or _
           LCase(leftTable.cell(i, 2).shape.TextFrame.textRange.text) Like "*falskt*" Then
            leftTable.Rows(i).Delete
        End If
    Next i

    For i = rightTable.Rows.count To 1 Step -1
        If LCase(rightTable.cell(i, 1).shape.TextFrame.textRange.text) Like "*false*" Or _
           LCase(rightTable.cell(i, 2).shape.TextFrame.textRange.text) Like "*false*" Or _
           LCase(rightTable.cell(i, 1).shape.TextFrame.textRange.text) Like "*falskt*" Or _
           LCase(rightTable.cell(i, 2).shape.TextFrame.textRange.text) Like "*falskt*" Then
            rightTable.Rows(i).Delete
        End If
    Next i
End Sub

Sub Abbrev_17()
    On Error Resume Next
    ActiveWindow.View.slide.Shapes("ImportedTable").Delete
End Sub

Sub Abbrev_Master()
    Abbrev_11
    Abbrev_12
    Abbrev_13
    Abbrev_14
    Abbrev_15
    Abbrev_16
    Abbrev_17
End Sub


