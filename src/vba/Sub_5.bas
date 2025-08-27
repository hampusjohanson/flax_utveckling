Attribute VB_Name = "Sub_5"
Sub Sub_5()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim sourceTable As table
    Dim rowIndex As Integer, colIndex As Integer
    Dim userName As String
    Dim filePath As String
    Dim operatingSystem As String

    ' Get the system username using Environ
    userName = Environ("USER") ' Get the username from the system environment

    ' Check if the username is empty
    If userName = "" Then
        MsgBox "No username found in system environment. Please provide a username.", vbCritical
        Exit Sub
    End If

    ' === Determine the operating system and file path ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' For macOS
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv" ' Use the username for macOS
    Else
        ' For Windows
        filePath = "C:\Users\" & userName & "\Desktop\exported_data_semi.csv" ' Build the file path for Windows
    End If

    ' === Check if the file exists ===
    If Dir(filePath) = "" Then
        MsgBox "The file was not found at: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Find the active slide and chart on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    Set chartShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape

    If chartShape Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Open chart data
    On Error Resume Next
    chartShape.chart.chartData.Activate
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    On Error GoTo 0

    If chartDataWorkbook Is Nothing Then
        MsgBox "Could not open chart data. Please check compatibility with macOS.", vbCritical
        Exit Sub
    End If

    ' Clear existing data in the range J2:Q52
    Set chartSheet = chartDataWorkbook.Worksheets(1)
    chartSheet.Range("J2:Q52").Clear

    ' Find the table with the name "table_substance"
    Set sourceTable = Nothing
    For Each shape In pptSlide.Shapes
        If shape.HasTable And shape.Name = "table_substance" Then
            Set sourceTable = shape.table
            Exit For
        End If
    Next shape

    If sourceTable Is Nothing Then
        MsgBox "The 'table_substance' table was not found on the slide.", vbExclamation
        chartDataWorkbook.Close
        Exit Sub
    End If

    ' Paste specific columns from the table into Excel
    ' Column 1 to J2 and down
    For rowIndex = 1 To sourceTable.Rows.count
        chartSheet.Cells(rowIndex + 1, 10).value = sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text
    Next rowIndex

    ' Column 2 to L2 and down
    For rowIndex = 1 To sourceTable.Rows.count
        chartSheet.Cells(rowIndex + 1, 12).value = sourceTable.cell(rowIndex, 2).shape.TextFrame.textRange.text
    Next rowIndex

    ' Column 5 to K2 and down
    For rowIndex = 1 To sourceTable.Rows.count
        chartSheet.Cells(rowIndex + 1, 11).value = sourceTable.cell(rowIndex, 5).shape.TextFrame.textRange.text
    Next rowIndex

    ' Columns 6 to 10 to M2 and down
    For rowIndex = 1 To sourceTable.Rows.count
        For colIndex = 6 To 10
            chartSheet.Cells(rowIndex + 1, colIndex + 7).value = sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text
        Next colIndex
    Next rowIndex

    ' Remove the table from the slide
    For Each shape In pptSlide.Shapes
        If shape.HasTable And shape.Name = "table_substance" Then
            shape.Delete
            Exit For
        End If
    Next shape

    ' Close the chart's data workbook
    chartDataWorkbook.Close
End Sub

