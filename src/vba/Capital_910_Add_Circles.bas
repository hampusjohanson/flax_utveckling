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

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' === Hitta den största tabellen ===
    maxCells = 0
    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            ' Räkna antalet celler (rader * kolumner)
            tableCells = tableShape.table.Rows.count * tableShape.table.Columns.count
            If tableCells > maxCells Then
                maxCells = tableCells
                Set largestTable = tableShape
            End If
        End If
    Next tableShape

    ' Kontrollera om en tabell hittades
    If largestTable Is Nothing Then
        MsgBox "Ingen tabell hittades på sliden.", vbCritical
        Exit Sub
    End If

    ' Namnge den största tabellen som "Cap_Table"
    largestTable.Name = "Cap_Table"

    ' Dynamiskt hämta filvägen till skrivbordet beroende på OS
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        ' För macOS
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv" ' Använd användarnamn för macOS
    Else
        ' För Windows
        filePath = "C:/Local\exported_data_semi.csv" ' Bygg filvägen för Windows
    End If

    ' Kontrollera om filen finns
    If Dir(filePath) = "" Then
        MsgBox "Filen hittades inte på: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Importera CSV och skapa SOURCE-tabellen ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Skapa en ny tabell på sliden (endast en kolumn för hexkoder)
    Set tableShape = pptSlide.Shapes.AddTable(numRows:=1, NumColumns:=1, left:=50, Top:=50, width:=100, height:=300)
    tableShape.Name = "SOURCE"
    Set sourceTable = tableShape.table

    ' Läs CSV och fyll tabellen (rad 21-40, endast kolumn 6)
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

            ' Kontrollera om kolumn 6 finns och att värdet inte är "false" eller "falskt"
            If UBound(Data) >= 5 Then
                If LCase(Trim(Data(5))) <> "false" And LCase(Trim(Data(5))) <> "falskt" Then
                    ' Lägg till en ny rad om det behövs
                    If currentRow > sourceTable.Rows.count Then
                        sourceTable.Rows.Add
                    End If

                    ' Fyll i cellen för hexkoden (endast kolumn 6)
                    sourceTable.cell(currentRow, 1).shape.TextFrame.textRange.text = Trim(Data(5))
                    currentRow = currentRow + 1
                End If
            End If
        End If
    Loop

    Close fileNumber

    ' Ta bort eventuellt tomma rader från tabellen
    For rowCounter = sourceTable.Rows.count To currentRow Step -1
        If Trim(sourceTable.cell(rowCounter, 1).shape.TextFrame.textRange.text) = "" Then
            sourceTable.Rows(rowCounter).Delete
        End If
    Next rowCounter

    ' === Skapa cirklar till vänster om Cap_Table ===
    leftPosition = largestTable.left - 20 ' Placera cirklar till vänster om tabellen
    rowHeight = largestTable.table.Rows(2).height ' Höjd på en rad (förutsatt att alla rader är lika)
    circleSize = rowHeight * 0.9 ' Cirklar ska vara 90% av radens höjd

    For rowIndex = 1 To largestTable.table.Rows.count - 1
        ' Kontrollera om det finns en motsvarande rad i SOURCE
        If rowIndex <= sourceTable.Rows.count Then
            ' Hämta hexkod från SOURCE-tabellen
            hexColor = Trim(sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text)
            If Len(hexColor) = 7 And left(hexColor, 1) = "#" Then
                r = CLng("&H" & Mid(hexColor, 2, 2))
                g = CLng("&H" & Mid(hexColor, 4, 2))
                b = CLng("&H" & Mid(hexColor, 6, 2))
            Else
                r = 200: g = 200: b = 200 ' Standardfärg om hexkoden är ogiltig
            End If
        Else
            r = 200: g = 200: b = 200 ' Om raden inte finns, använd standardfärg
        End If

        ' Beräkna topposition för cirkeln
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

    ' Ta bort SOURCE-tabellen efter användning
    tableShape.Delete

End Sub

