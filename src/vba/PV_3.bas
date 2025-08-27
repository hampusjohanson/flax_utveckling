Attribute VB_Name = "PV_3"
Sub PV_3()
    ' Variabler
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim tableShape As shape
    Dim sourceTable As table
    Dim csvData(419 To 428, 1 To 2) As String
    Dim operatingSystem As String
    Dim userName As String

    ' === Kontrollera och hämta användarnamn från Environ ===
    userName = Environ("USER") ' Get the username directly from the system environment

    ' Dynamiskt skapa filväg baserat på användarnamnet och operativsystem
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Kontrollera om filen finns
    If Dir(filePath) = "" Then
        MsgBox "Filen hittades inte på: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Läs in data från CSV ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    rowIndex = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        If rowIndex >= 419 And rowIndex <= 428 Then
            Data = Split(line, ";")
            csvData(rowIndex, 2) = Trim(Data(1)) ' Kolumn 2
            If UBound(Data) >= 2 Then
                csvData(rowIndex, 1) = Trim(Data(2)) ' Kolumn 3
            Else
                csvData(rowIndex, 1) = ""
            End If
        End If
    Loop

    Close fileNumber

    ' === Hitta tabellen på sliden ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            Set sourceTable = tableShape.table
            Exit For
        End If
    Next tableShape

    If sourceTable Is Nothing Then
        MsgBox "Ingen tabell hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' === Klistra in värden från CSV till tabellen ===
    For rowIndex = 1 To 50
        If rowIndex + 1 <= sourceTable.Rows.count Then
            ' Kolumn 2 från CSV -> Kolumn 5 i tabellen
            sourceTable.cell(rowIndex + 1, 5).shape.TextFrame.textRange.text = csvData(rowIndex + 418, 2)

            ' Kolumn 3 från CSV -> Kolumn 2 i tabellen
            sourceTable.cell(rowIndex + 1, 2).shape.TextFrame.textRange.text = csvData(rowIndex + 418, 1)
        End If
    Next rowIndex

End Sub

