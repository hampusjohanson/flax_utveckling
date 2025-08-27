Attribute VB_Name = "Export_PDF"
Sub ExportToPDF()
    Dim pdfPath As String
    Dim defaultPath As String
    Dim fileName As String
    Dim datePrefix As String

    ' Set the date prefix
    datePrefix = Format(Date, "yyyy-mm-dd") & " "

    ' Get the current file name without extension
    fileName = Replace(ActivePresentation.Name, ".pptx", "")
    fileName = Replace(fileName, ".pptm", "")

    ' Check if the file name already contains a date-like text
    If Not fileName Like "####-##-##*" Then
        fileName = datePrefix & fileName
    End If

    ' Set default path to the current file directory
    If ActivePresentation.Path <> "" Then
        defaultPath = ActivePresentation.Path & "\\" & fileName & ".pdf"
    Else
        defaultPath = fileName & ".pdf" ' Default name if file is unsaved
    End If

    ' Prompt user for file path
    pdfPath = InputBox("Enter the path to save the PDF:", "Export to PDF", defaultPath)

    ' Check if the user entered a path
    If pdfPath <> "" Then
        On Error Resume Next
        ActivePresentation.ExportAsFixedFormat pdfPath, ppFixedFormatTypePDF, ppFixedFormatIntentPrint
        If Err.Number = 0 Then
            MsgBox "PDF successfully exported to: " & pdfPath, vbInformation, "Export Complete"
        Else
            MsgBox "An error occurred while exporting to PDF.", vbCritical, "Export Failed"
        End If
        On Error GoTo 0
    Else
        MsgBox "Export canceled.", vbExclamation, "Export Canceled"
    End If
End Sub

