Attribute VB_Name = "Lines_5"
Sub Lines_5()
    ' F�rgs�tter linjer i diagram baserat p� hex-koder fr�n CSV (tyst version)
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim hexColors(1 To 12) As String
    Dim red As Integer, green As Integer, blue As Integer
    Dim operatingSystem As String
    Dim userName As String
    Dim rowIndex As Integer
    Dim SeriesIndex As Integer
    Dim hexCodeRowStart As Integer
    Dim hexCodeRowEnd As Integer
    Dim totalSeries As Integer
    Dim maxSeries As Integer

    ' === Filv�g ===
    userName = Environ("USER")
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If
    Debug.Print "File path: " & filePath

    If Dir(filePath) = "" Then Exit Sub

    ' === L�s in f�rger ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    rowIndex = 1
    SeriesIndex = 1
    hexCodeRowStart = 837
    hexCodeRowEnd = hexCodeRowStart + 11

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        If rowIndex >= hexCodeRowStart And rowIndex <= hexCodeRowEnd Then
            Data = Split(line, ";")
            If UBound(Data) >= 0 Then
                hexColors(SeriesIndex) = Trim(Replace(Replace(Data(0), vbCr, ""), vbLf, ""))
                If hexColors(SeriesIndex) = "" Then
                    Debug.Print "? Hex-kod saknas p� rad " & rowIndex
                    Close fileNumber: Exit Sub
                ElseIf left(hexColors(SeriesIndex), 1) <> "#" Or Len(hexColors(SeriesIndex)) <> 7 Then
                    Debug.Print "? Ogiltig hex-kod p� rad " & rowIndex & ": " & hexColors(SeriesIndex)
                    Close fileNumber: Exit Sub
                End If
                Debug.Print "? F�rg " & SeriesIndex & ": " & hexColors(SeriesIndex)
                SeriesIndex = SeriesIndex + 1
            Else
                Debug.Print "? Rad " & rowIndex & " saknar kolumn 1."
                Close fileNumber: Exit Sub
            End If
        End If
        rowIndex = rowIndex + 1
        If SeriesIndex > 12 Then Exit Do
    Loop
    Close fileNumber

    ' === G� igenom diagram p� aktuell slide ===
    Set pptSlide = ActiveWindow.View.slide
    For Each chartShape In pptSlide.Shapes
        If chartShape.Type = msoChart Then
            Set chartObject = chartShape.chart
            totalSeries = chartObject.SeriesCollection.count
            Debug.Print "?? Diagram: " & chartShape.Name & " har " & totalSeries & " serier"

            If totalSeries < 12 Then
                Debug.Print "? Diagram """ & chartShape.Name & """ har bara " & totalSeries & " serier (minst 12 kr�vs)."
                GoTo N�staDiagram
            End If

            maxSeries = IIf(totalSeries < 12, totalSeries, 12)

            For SeriesIndex = 1 To maxSeries
                Set series = chartObject.SeriesCollection(SeriesIndex)
                On Error GoTo F�rgfel
                red = CLng("&H" & Mid(hexColors(SeriesIndex), 2, 2))
                green = CLng("&H" & Mid(hexColors(SeriesIndex), 4, 2))
                blue = CLng("&H" & Mid(hexColors(SeriesIndex), 6, 2))
                series.Format.line.ForeColor.RGB = RGB(red, green, blue)
                Debug.Print "?? Serie " & SeriesIndex & " f�rgsatt: RGB(" & red & "," & green & "," & blue & ")"
                On Error GoTo 0
            Next SeriesIndex
        Else
            Debug.Print "? Hoppar �ver shape: " & chartShape.Name & " (ej diagram)"
        End If
N�staDiagram:
    Next chartShape
    Exit Sub

F�rgfel:
    Debug.Print "? Fel vid f�rgs�ttning av serie " & SeriesIndex & ": " & hexColors(SeriesIndex)
    Resume Next
End Sub

