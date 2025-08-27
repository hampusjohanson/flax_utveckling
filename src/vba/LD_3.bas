Attribute VB_Name = "LD_3"
Sub Mac_LD_1()
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
    Dim csvData(93 To 102, 1 To 2) As String
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

        If rowIndex >= 93 And rowIndex <= 102 Then
            Data = Split(line, ";")
            csvData(rowIndex, 1) = Trim(Data(0)) ' Kolumn A
            If UBound(Data) >= 1 Then
                csvData(rowIndex, 2) = Trim(Data(1)) ' Kolumn B
            Else
                csvData(rowIndex, 2) = ""
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
            ' Kolumn A från CSV -> Kolumn 2 i tabellen
            sourceTable.cell(rowIndex + 1, 2).shape.TextFrame.textRange.text = csvData(rowIndex + 92, 1)

            ' Kolumn B från CSV -> Kolumn 4 i tabellen
            sourceTable.cell(rowIndex + 1, 5).shape.TextFrame.textRange.text = csvData(rowIndex + 92, 2)
        End If
    Next rowIndex

End Sub

