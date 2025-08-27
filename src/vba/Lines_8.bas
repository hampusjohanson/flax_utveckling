Attribute VB_Name = "Lines_8"
' CleanText Function (removes non-printable characters, accents, and other invisible characters)
Function cleanText(inputText As String) As String
    Dim cleaned As String
    Dim normalizedText As String
    cleaned = Trim(inputText) ' Remove leading/trailing spaces
    cleaned = Replace(cleaned, Chr(160), "") ' Remove non-breaking spaces (CHAR(160))
    
    ' Remove line breaks and other non-printable characters
    cleaned = Replace(cleaned, Chr(10), "") ' Remove line breaks
    cleaned = Replace(cleaned, Chr(13), "") ' Remove carriage returns
    cleaned = Replace(cleaned, Chr(9), "") ' Remove tab characters
    
    ' Remove diacritics (accents) using VBA's CreateObject
    normalizedText = RemoveAccents(cleaned)

    cleanText = normalizedText
End Function

' Function to remove accents from characters
Function RemoveAccents(inputText As String) As String
    Dim i As Integer
    Dim result As String
    Dim accents As String
    Dim unaccented As String

    ' Define accented characters and their unaccented equivalents
    accents = "áéíóúÁÉÍÓÚàèìòùÀÈÌÒÙäëïöüÄËÏÖÜâêîôûÂÊÎÔÛ"
    unaccented = "aeiouAEIOUaeiouAEIOUaeiouAEIOUaeiouAEIOU"

    result = inputText

    ' Loop through each character in the input and replace accents with unaccented characters
    For i = 1 To Len(accents)
        result = Replace(result, Mid(accents, i, 1), Mid(unaccented, i, 1))
    Next i

    RemoveAccents = result
End Function


Sub SetMarkersForAllP()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim rowData As Object
    Dim seriesNames As Object
    Dim pValue As String
    Dim seriesName As String
    Dim MatchFound As Boolean
    Dim SeriesIndex As Integer
    Dim rowIndex As Integer
    Dim lineColor As Long

    ' === Loop through all charts on the slide ===
    Set pptSlide = ActiveWindow.View.slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart

            ' Access the embedded Excel workbook and worksheet
            Set excelWorkbook = chartObject.chartData.Workbook
            Set excelWorksheet = excelWorkbook.Sheets(1) ' The embedded Excel worksheet

            ' Access column P (P2:P51) and series names in D1:O1 (12 brands)
            Set seriesNames = excelWorksheet.Range("D1:O1")

            ' Loop through P2:P51 to check each P value
            For rowIndex = 2 To 51
                Set rowData = excelWorksheet.Range("P" & rowIndex)
                pValue = cleanText(rowData.value)

                Debug.Print "P" & rowIndex & " Value: " & pValue

                MatchFound = False
                For SeriesIndex = 1 To seriesNames.Cells.count
                    seriesName = cleanText(seriesNames.Cells(1, SeriesIndex).value)
                    Debug.Print "Comparing P" & rowIndex & " (" & pValue & ") with Series name (" & seriesName & ")"

                    If LCase(pValue) = LCase(seriesName) Then
                        MatchFound = True
                        Debug.Print "P" & rowIndex & " matches Series name: " & seriesName

                        If chartObject.SeriesCollection(SeriesIndex).Format.line.visible = msoTrue And _
                           chartObject.SeriesCollection(SeriesIndex).MarkerStyle <> msoMarkerNone Then

                            lineColor = chartObject.SeriesCollection(SeriesIndex).Format.line.ForeColor.RGB

                            chartObject.SeriesCollection(SeriesIndex).Points(rowIndex - 1).MarkerStyle = xlMarkerStyleCircle
                            chartObject.SeriesCollection(SeriesIndex).Points(rowIndex - 1).MarkerSize = 6
                            chartObject.SeriesCollection(SeriesIndex).Points(rowIndex - 1).MarkerBackgroundColor = lineColor
                        End If

                        Exit For
                    End If
                Next SeriesIndex

                If Not MatchFound Then
                    Debug.Print "No match found for P" & rowIndex & " in series names."
                End If
            Next rowIndex
        End If
    Next chartShape
End Sub


