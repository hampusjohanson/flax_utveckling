Attribute VB_Name = "Table_Tool_2"
Sub Table_Tool_2()
    Dim pptSlide As slide
    Dim selectedTable As table
    Dim shapeItem As shape
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim operatingSystem As String
    Dim userName As String
    Dim metricColumn As Integer
    Dim rowIndex As Integer
    Dim currentRow As Integer
    Dim rowLimit As Integer
    Dim foundMetric As Boolean
    
    ' === Check if a table is selected ===
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    Set shapeItem = ActiveWindow.Selection.ShapeRange(1)
    If Not shapeItem.HasTable Then Exit Sub

    Set selectedTable = shapeItem.table

    ' === Exit if no valid metric or row count ===
    If metric_decision = "" Or row_decision < 1 Then Exit Sub

    ' === Determine file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        userName = Environ("USER")
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        userName = Environ("USERNAME")
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' === Check if file exists ===
    If Dir(filePath) = "" Then Exit Sub

    ' === Open the CSV file ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    currentRow = 0
    foundMetric = False

    ' === Find column for metric decision (Row 418) ===
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        currentRow = currentRow + 1

        If currentRow = 418 Then
            Data = Split(line, ";")
            For metricColumn = 0 To UBound(Data)
                If Trim(Data(metricColumn)) = metric_decision Then
                    foundMetric = True
                    Exit For
                End If
            Next metricColumn
            Exit Do
        End If
    Loop

    ' === If metric not found, exit ===
    If Not foundMetric Then Close fileNumber: Exit Sub

    ' === Insert values from Row 419 and down ===
    rowIndex = 2 ' Start at Row 2 in the selected table
    rowLimit = row_decision ' Number of rows to insert
    currentRow = 418 ' Start tracking row index

    Do While Not EOF(fileNumber) And rowIndex <= rowLimit + 1 ' Adjust the condition here
        Line Input #fileNumber, line
        currentRow = currentRow + 1

        ' Start inserting from row 419
        If currentRow >= 419 Then
            Data = Split(line, ";")
            If UBound(Data) >= metricColumn Then
                ' === Expand table if needed ===
                If rowIndex > selectedTable.Rows.count Then selectedTable.Rows.Add

                ' === Insert into Column 2 ===
                selectedTable.cell(rowIndex, 2).shape.TextFrame.textRange.text = Trim(Data(metricColumn))
                rowIndex = rowIndex + 1
            End If
        End If
    Loop

    Close fileNumber

    ' === Remove extra rows ===
    While selectedTable.Rows.count > rowLimit + 1
        selectedTable.Rows(selectedTable.Rows.count).Delete
    Wend
End Sub

