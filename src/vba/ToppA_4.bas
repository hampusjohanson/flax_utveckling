Attribute VB_Name = "ToppA_4"
Sub ToppA_4()
    ' Define variables
    Dim csvFilePath As String
    Dim userName As String
    Dim operatingSystem As String
    Dim csvContent As String
    Dim csvLines() As String
    Dim startRow As Long
    Dim endRow As Long
    Dim currentSlide As slide
    Dim longStrongerTable As table
    Dim longWeakerTable As table
    Dim rowIndex As Long
    Dim sourceRow As Integer

    ' Get the username and operating system
    userName = Environ("USER")
    operatingSystem = Application.operatingSystem

    ' Define file path based on OS
    If InStr(operatingSystem, "Macintosh") > 0 Then
        csvFilePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        csvFilePath = "C:\\Local\\exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(csvFilePath) = "" Then
        MsgBox "File not found: " & csvFilePath
        Exit Sub
    End If

    ' Read the CSV file
    Open csvFilePath For Input As #1
    csvContent = Input(LOF(1), #1)
    Close #1

    ' Split CSV into lines
    csvLines = Split(csvContent, vbLf)

    ' Define start and end rows
    startRow = 684
    endRow = 733

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' Dimensions for table
    Dim width As Single
    Dim height As Single
    width = 15 * 28.35
    height = 0.56 * 28.35

    ' Create the Stronger table
    On Error GoTo ErrorHandler
    Set shapeStronger = currentSlide.Shapes.AddTable(1, 1, 8.67 * 28.35, 5.76 * 28.35, width, height)
    shapeStronger.Name = "long_stronger"
    Set longStrongerTable = shapeStronger.table
    On Error GoTo 0 ' Disable error handling

    ' Create the Weaker table
    Set shapeWeaker = currentSlide.Shapes.AddTable(1, 1, 24.45 * 28.35, 5.76 * 28.35, width, height)
    shapeWeaker.Name = "long_weaker"
    Set longWeakerTable = shapeWeaker.table

    ' Populate the long_stronger table
    sourceRow = 1
    For rowIndex = startRow To endRow
        If rowIndex - 1 > UBound(csvLines) Then
            Debug.Print "Row " & rowIndex & " exceeds the file's content."
            Exit Sub
        End If

        Dim Columns() As String
        Columns = Split(csvLines(rowIndex - 1), ";")

        ' Ensure column values are trimmed and skip invalid rows
        If UBound(Columns) >= 3 Then
            Dim col4 As String, col3 As String
            col4 = Trim(Columns(3))
            col3 = Trim(Columns(2))

            ' Skip rows with "False" values
            If col4 = "False" Or col3 = "False" Then
                Debug.Print "Skipping row " & rowIndex & " due to invalid data (False values)."
                GoTo ContinueLoop
            End If

            ' Populate long_stronger if column 3 = 1
            If col3 = "1" Then
                If sourceRow > longStrongerTable.Rows.count Then
                    longStrongerTable.Rows.Add
                    Debug.Print "Added row to long_stronger table at row: " & sourceRow
                End If
                longStrongerTable.cell(sourceRow, 1).shape.TextFrame.textRange.text = col4
                Debug.Print "Row " & sourceRow & " populated with: " & col4
                sourceRow = sourceRow + 1
            End If
        End If
ContinueLoop:
    Next rowIndex

    ' Reset sourceRow and process rows for long_weaker (column 3 = 2)
    sourceRow = 1
    For rowIndex = startRow To endRow
        If rowIndex - 1 > UBound(csvLines) Then
            Debug.Print "Row " & rowIndex & " exceeds the file's content."
            Exit Sub
        End If

        Dim weakerColumns() As String
        weakerColumns = Split(csvLines(rowIndex - 1), ";")

        ' Ensure column values are trimmed and skip invalid rows
        If UBound(weakerColumns) >= 3 Then
            Dim weakerCol4 As String, weakerCol3 As String
            weakerCol4 = Trim(weakerColumns(3))
            weakerCol3 = Trim(weakerColumns(2))

            ' Skip rows with "False" values
            If weakerCol4 = "False" Or weakerCol3 = "False" Then
                Debug.Print "Skipping row " & rowIndex & " due to invalid data (False values)."
                GoTo ContinueLoopWeaker
            End If

            ' Populate long_weaker if column 3 = 2
            If weakerCol3 = "2" Then
                If sourceRow > longWeakerTable.Rows.count Then
                    longWeakerTable.Rows.Add
                    Debug.Print "Added row to long_weaker table at row: " & sourceRow
                End If
                longWeakerTable.cell(sourceRow, 1).shape.TextFrame.textRange.text = weakerCol4
                Debug.Print "Row " & sourceRow & " populated with: " & weakerCol4
                sourceRow = sourceRow + 1
            End If
        End If
ContinueLoopWeaker:
    Next rowIndex

    Debug.Print "Tables populated successfully."
    Exit Sub ' Exit after table creation

ErrorHandler:
    MsgBox "Error during table creation: " & Err.Description
End Sub

