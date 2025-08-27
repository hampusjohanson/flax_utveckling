Attribute VB_Name = "Module5"
Sub WriteFormulaDebug()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartDataWorkbook As Object
    Dim chartSheet As Object
    Dim formula As String
    Dim delimiter As String

    ' Validate that a language is selected
    If SelectedLanguage = "" Then
        MsgBox "Please select a language from the Ribbon first.", vbExclamation
        Exit Sub
    End If

    ' Set the delimiter manually based on regional settings
    delimiter = ";"

    ' Get the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart on the current slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Exit For
        End If
    Next chartShape

    ' If no chart is found, exit the macro
    If chartShape Is Nothing Then
        MsgBox "No chart found on the current slide.", vbCritical
        Exit Sub
    End If

    ' Open the chart's Excel workbook
    Set chartDataWorkbook = chartShape.chart.chartData.Workbook
    Set chartSheet = chartDataWorkbook.Sheets(1) ' Access the first sheet

    ' Construct the formula based on the selected language
    If SelectedLanguage = "Swedish" Then
        formula = "=OM(ÄRTEXT(L2)" & delimiter & "VÄRDE(L2)" & delimiter & "FALSKT)"
    ElseIf SelectedLanguage = "English" Then
        formula = "=IF(ISTEXT(L2)" & delimiter & "VALUE(L2)" & delimiter & "FALSE)"
    Else
        MsgBox "Unknown language selected.", vbCritical
        Exit Sub
    End If

    ' Debugging: Print the formula to verify it
    Debug.Print "Formula: " & formula

    ' Attempt to write the formula
    On Error Resume Next
    chartSheet.Range("B2").formula = formula
    If Err.Number <> 0 Then
        MsgBox "Error writing the formula to B2: " & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Close the workbook to save changes
    chartShape.chart.chartData.Workbook.Close

    MsgBox "Formula written to B2 in the chart's Excel workbook!", vbInformation
End Sub


