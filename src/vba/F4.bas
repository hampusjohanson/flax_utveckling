Attribute VB_Name = "F4"
Sub Abb_1()
    Dim pptSlide As slide
    Dim userName As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim importedTableShape As shape
    Dim sourceTable As table
    Dim startRow As Integer, endRow As Integer
    Dim startCol As Integer, endCol As Integer
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellValue As String
    Dim tableRow As Integer, tableCol As Integer
    Dim operatingSystem As String

    ' === Kontrollera operativsystem ===
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' F�r macOS
        Set pptSlide = ActivePresentation.Slides(1)
        If pptSlide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
            userName = Trim(Split(pptSlide.NotesPage.Shapes.Placeholders(2).TextFrame.textRange.text, vbCrLf)(0))
        Else
            MsgBox "Speaker Notes p� Slide 1 �r tomma. Ange ditt anv�ndarnamn p� f�rsta raden.", vbCritical
            Exit Sub
        End If
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        ' F�r Windows
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Kontrollera om filen finns
    If Dir(filePath) = "" Then
        MsgBox "Filen hittades inte p�: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === H�rdkodade rader och kolumner (ange v�rden h�r) ===
    startRow = 392  ' B�rja fr�n rad 392
    endRow = 417    ' Sluta p� rad 417
    startCol = 1    ' B�rja fr�n kolumn 1
    endCol = 5      ' Sluta p� kolumn 5

    ' �ppna filen f�r att l�sa data
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Skapa en ny tabell p� den aktuella sliden
    Set pptSlide = ActiveWindow.View.slide
    Set importedTableShape = pptSlide.Shapes.AddTable(numRows:=endRow - startRow + 1, NumColumns:=endCol - startCol + 1, left:=50, Top:=50, width:=600, height:=300)
    Set sourceTable = importedTableShape.table

    ' === L�s och fyll hela tabellen ===
    rowIndex = 0
    tableRow = 1
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' Om raden �r inom det h�rdkodade intervallet
        If rowIndex >= startRow And rowIndex <= endRow Then
            Data = Split(line, ";")
            tableCol = 1
            For colIndex = startCol To endCol
                If colIndex <= UBound(Data) + 1 Then
                    cellValue = Trim(Data(colIndex - 1))
                    sourceTable.cell(tableRow, tableCol).shape.TextFrame.textRange.text = cellValue
                    tableCol = tableCol + 1
                End If
            Next colIndex
            tableRow = tableRow + 1
        End If
    Loop

    ' St�ng filen
    Close fileNumber

    ' === Rensa rader baserat p� kolumn 4 ===
    For rowIndex = sourceTable.Rows.count To 1 Step -1
        cellValue = LCase(Trim(sourceTable.cell(rowIndex, 4).shape.TextFrame.textRange.text))
        If cellValue = "false" Or cellValue = "falskt" Then
            sourceTable.Rows(rowIndex).Delete
        End If
    Next rowIndex

    ' === Rensa alla celler med "false" eller "falskt" oavsett kolumn ===
    For rowIndex = 1 To sourceTable.Rows.count
        For colIndex = 1 To sourceTable.Columns.count
            cellValue = LCase(Trim(sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text))
            If cellValue = "false" Or cellValue = "falskt" Then
                sourceTable.cell(rowIndex, colIndex).shape.TextFrame.textRange.text = ""
            End If
        Next colIndex
    Next rowIndex

    MsgBox "Data importerades och rensades framg�ngsrikt fr�n: " & filePath, vbInformation
End Sub

