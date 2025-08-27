Attribute VB_Name = "Capital_910_Add_Circles"
Sub Mac_Cap_Circles()
    Dim pptSlide As slide
    Dim largestTable As shape
    Dim tableShape As shape
    Dim sourceTable As table
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim currentRow As Integer
    Dim Data() As String
    Dim rowIndex As Integer
    Dim circleShape As shape
    Dim leftPosition As Double
    Dim topPosition As Double
    Dim rowHeight As Double
    Dim circleSize As Double
    Dim hexColor As String
    Dim r As Integer, g As Integer, b As Integer
    Dim maxCells As Integer
    Dim tableCells As Integer
    Dim operatingSystem As String
    Dim userName As String

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' === Hitta den st�rsta tabellen ===
    maxCells = 0
    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            ' R�kna antalet celler (rader * kolumner)
            tableCells = tableShape.table.Rows.count * tableShape.table.Columns.count
            If tableCells > maxCells Then
                maxCells = tableCells
                Set largestTable = tableShape
            End If
        End If
    Next tableShape

    ' Kontrollera om en tabell hittades
    If largestTable Is Nothing Then
        MsgBox "Ingen tabell hittades p� sliden.", vbCritical
        Exit Sub
    End If

    ' Namnge den st�rsta tabellen som "Cap_Table"
    largestTable.Name = "Cap_Table"

    ' Dynamiskt h�mta filv�gen till skrivbordet beroende p� OS
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' F�r macOS
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv" ' Anv�nd anv�ndarnamn f�r macOS
    Else
        ' F�r Windows
        filePath = "C:/Local\exported_data_semi.csv" ' Bygg filv�gen f�r Windows
    End If

    ' Kontrollera om filen finns
    If Dir(filePath) = "" Then
        MsgBox "Filen hittades inte p�: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Importera CSV och skapa SOURCE-tabellen ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Skapa en ny tabell p� sliden (endast en kolumn f�r hexkoder)
    Set tableShape = pptSlide.Shapes.AddTable(numRows:=1, NumColumns:=1, left:=50, Top:=50, width:=100, height:=300)
    tableShape.Name = "SOURCE"
    Set sourceTable = tableShape.table

    ' L�s CSV och fyll tabellen (rad 21-40, endast kolumn 6)
    currentRow = 1
    Dim startRow As Integer: startRow = 21
    Dim endRow As Integer: endRow = 40
    Dim rowCounter As Integer: rowCounter = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowCounter = rowCounter + 1

        ' Bearbeta endast rader mellan startRow och endRow
        If rowCounter >= startRow And rowCounter <= endRow Then
            Data = Split(line, ";")

            ' Kontrollera om kolumn 6 finns och att v�rdet inte �r "false" eller "falskt"
            If UBound(Data) >= 5 Then
                If LCase(Trim(Data(5))) <> "false" And LCase(Trim(Data(5))) <> "falskt" Then
                    ' L�gg till en ny rad om det beh�vs
                    If currentRow > sourceTable.Rows.count Then
                        sourceTable.Rows.Add
                    End If

                    ' Fyll i cellen f�r hexkoden (endast kolumn 6)
                    sourceTable.cell(currentRow, 1).shape.TextFrame.textRange.text = Trim(Data(5))
                    currentRow = currentRow + 1
                End If
            End If
        End If
    Loop

    Close fileNumber

    ' Ta bort eventuellt tomma rader fr�n tabellen
    For rowCounter = sourceTable.Rows.count To currentRow Step -1
        If Trim(sourceTable.cell(rowCounter, 1).shape.TextFrame.textRange.text) = "" Then
            sourceTable.Rows(rowCounter).Delete
        End If
    Next rowCounter

    ' === Skapa cirklar till v�nster om Cap_Table ===
    leftPosition = largestTable.left - 20 ' Placera cirklar till v�nster om tabellen
    rowHeight = largestTable.table.Rows(2).height ' H�jd p� en rad (f�rutsatt att alla rader �r lika)
    circleSize = rowHeight * 0.9 ' Cirklar ska vara 90% av radens h�jd

    For rowIndex = 1 To largestTable.table.Rows.count - 1
        ' Kontrollera om det finns en motsvarande rad i SOURCE
        If rowIndex <= sourceTable.Rows.count Then
            ' H�mta hexkod fr�n SOURCE-tabellen
            hexColor = Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)
            If Len(hexColor) = 7 And left(hexColor, 1) = "#" Then
                r = CLng("&H" & Mid(hexColor, 2, 2))
                g = CLng("&H" & Mid(hexColor, 4, 2))
                b = CLng("&H" & Mid(hexColor, 6, 2))
            Else
                r = 200: g = 200: b = 200 ' Standardf�rg om hexkoden �r ogiltig
            End If
        Else
            r = 200: g = 200: b = 200 ' Om raden inte finns, anv�nd standardf�rg
        End If

        ' Ber�kna topposition f�r cirkeln
        topPosition = largestTable.Top + largestTable.table.Rows(1).height + (rowIndex - 1) * rowHeight

        ' Skapa cirkel
        Set circleShape = pptSlide.Shapes.AddShape(msoShapeOval, leftPosition, topPosition, circleSize, circleSize)
        With circleShape
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(r, g, b)
            .line.visible = msoFalse ' Ingen kantlinje
            .Name = "Circle" & rowIndex ' Namnge cirkeln
        End With
    Next rowIndex

    ' Ta bort SOURCE-tabellen efter anv�ndning
    tableShape.Delete

End Sub

