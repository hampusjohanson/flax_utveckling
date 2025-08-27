Attribute VB_Name = "ToppA_1"
Sub ToppA_1()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you on that drivkrafts-slide with TWO COLUMNS NEXT TO EACH OTHER?", vbYesNo + vbQuestion, "Check Slide")
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbExclamation
        Exit Sub
    End If

    Dim csvFilePath As String
    Dim userName As String
    Dim operatingSystem As String
    Dim csvContent As String
    Dim csvLines() As String
    Dim startRow As Long
    Dim endRow As Long
    Dim currentSlide As slide
    Dim pptTable As table
    Dim strongerTable As table
    Dim weakerTable As table
    Dim rowIndex As Long
    Dim sourceRow As Integer
    Dim targetRow As Integer

    ' Get the username from the Environ function
    userName = Environ("USER")

    ' Determine the operating system and file path
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        csvFilePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        csvFilePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(csvFilePath) = "" Then
        MsgBox "File not found at: " & csvFilePath, vbExclamation
        Exit Sub
    End If

    ' Read the CSV file
    Open csvFilePath For Input As #1
    csvContent = Input(LOF(1), #1)
    Close #1

    csvContent = Replace(csvContent, vbCrLf, vbLf)
    csvLines = Split(csvContent, vbLf)

    startRow = 684
    endRow = 733

    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)
    Set pptTable = currentSlide.Shapes.AddTable(1, 2).table

    sourceRow = 1
    For rowIndex = startRow To endRow
        If rowIndex - 1 > UBound(csvLines) Then Exit Sub

        Dim Columns() As String
        Columns = Split(csvLines(rowIndex - 1), ";")

        If UBound(Columns) >= 2 Then
            Dim col1 As String, col2 As String, col3 As String
            col1 = Trim(Columns(0))
            col2 = Trim(Columns(1))
            col3 = Trim(Columns(2))

            If col3 = "1" Then
                If sourceRow > pptTable.Rows.count Then pptTable.Rows.Add
                pptTable.cell(sourceRow, 1).shape.TextFrame.textRange.text = col1
                pptTable.cell(sourceRow, 2).shape.TextFrame.textRange.text = col2
                sourceRow = sourceRow + 1
            End If
        End If
    Next rowIndex

    Dim shape As shape
    Dim tableCount As Integer
    tableCount = 0

    For Each shape In currentSlide.Shapes
        If shape.HasTable Then
            tableCount = tableCount + 1
            If tableCount = 1 Then
                shape.Name = "Stronger"
            ElseIf tableCount = 2 Then
                shape.Name = "Weaker"
            End If
        End If
    Next shape

    If tableCount < 2 Then
        MsgBox "Less than two tables found on the slide.", vbExclamation
    End If

    On Error Resume Next
    Set strongerTable = currentSlide.Shapes("Stronger").table
    On Error GoTo 0

    If strongerTable Is Nothing Then
        MsgBox "The table named 'Stronger' was not found.", vbExclamation
        Exit Sub
    End If

    targetRow = 2
    For sourceRow = 1 To pptTable.Rows.count
        If targetRow > strongerTable.Rows.count Then strongerTable.Rows.Add
        strongerTable.cell(targetRow, 1).shape.TextFrame.textRange.text = pptTable.cell(sourceRow, 1).shape.TextFrame.textRange.text
        strongerTable.cell(targetRow, 2).shape.TextFrame.textRange.text = pptTable.cell(sourceRow, 2).shape.TextFrame.textRange.text
        targetRow = targetRow + 1
    Next sourceRow

    Dim lastRow As Integer
    lastRow = targetRow - 1

    Do While strongerTable.Rows.count > lastRow
        strongerTable.Rows(strongerTable.Rows.count).Delete
    Loop

    sourceRow = 1
    For rowIndex = startRow To endRow
        If rowIndex - 1 > UBound(csvLines) Then Exit Sub

        Dim weakerColumns() As String
        weakerColumns = Split(csvLines(rowIndex - 1), ";")

        If UBound(weakerColumns) >= 2 Then
            Dim weakerCol1 As String, weakerCol2 As String, weakerCol3 As String
            weakerCol1 = Trim(weakerColumns(0))
            weakerCol2 = Trim(weakerColumns(1))
            weakerCol3 = Trim(weakerColumns(2))

            If weakerCol3 = "2" Then
                If sourceRow > pptTable.Rows.count Then pptTable.Rows.Add
                pptTable.cell(sourceRow, 1).shape.TextFrame.textRange.text = weakerCol1
                pptTable.cell(sourceRow, 2).shape.TextFrame.textRange.text = weakerCol2
                sourceRow = sourceRow + 1
            End If
        End If
    Next rowIndex

    On Error Resume Next
    Set weakerTable = currentSlide.Shapes("Weaker").table
    On Error GoTo 0

    If weakerTable Is Nothing Then
        MsgBox "The table named 'Weaker' was not found.", vbExclamation
        Exit Sub
    End If

    targetRow = 2
    For sourceRow = 1 To pptTable.Rows.count
        If targetRow > weakerTable.Rows.count Then weakerTable.Rows.Add
        weakerTable.cell(targetRow, 1).shape.TextFrame.textRange.text = pptTable.cell(sourceRow, 1).shape.TextFrame.textRange.text
        weakerTable.cell(targetRow, 2).shape.TextFrame.textRange.text = pptTable.cell(sourceRow, 2).shape.TextFrame.textRange.text
        targetRow = targetRow + 1
    Next sourceRow

    lastRow = targetRow - 1

    Do While weakerTable.Rows.count > lastRow
        weakerTable.Rows(weakerTable.Rows.count).Delete
    Loop

    currentSlide.Shapes(pptTable.Parent.Name).Delete
End Sub

