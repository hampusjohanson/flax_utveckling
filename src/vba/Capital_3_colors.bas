Attribute VB_Name = "Capital_3_colors"
Sub Mac_Cap_color()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim sourceTable As table
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim currentRow As Long
    Dim Data() As String
    Dim rowIndex As Integer
    Dim hexColor As String
    Dim r As Integer, g As Integer, b As Integer
    Dim operatingSystem As String
    Dim userName As String

    ' Error handling
    On Error GoTo ErrorHandler

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Hitta diagrammet på sliden
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' Dynamiskt hämta filvägen till skrivbordet beroende på OS
    operatingSystem = Application.operatingSystem
    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Kontrollera om filen finns
    If Dir(filePath) = "" Then
        MsgBox "Filen hittades inte på: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Importera CSV och skapa SOURCE-tabellen ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber

    ' Skapa en tillfällig tabell för hexkoder (endast en kolumn)
    Dim tableShape As shape
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

        If rowCounter >= startRow And rowCounter <= endRow Then
            Data = Split(line, ";")

            If UBound(Data) >= 5 Then
                If LCase(Trim(Data(5))) <> "false" And LCase(Trim(Data(5))) <> "falskt" And Trim(Data(5)) <> "#FFFFFF" Then
                    If currentRow > sourceTable.Rows.count Then
                        sourceTable.Rows.Add
                    End If

                    sourceTable.cell(currentRow, 1).shape.TextFrame.textRange.text = Trim(Data(5))
                    currentRow = currentRow + 1
                End If
            End If
        End If
    Loop

    Close fileNumber

    ' === Färglägg datapunkter i diagrammet ===
    With chartShape.chart.SeriesCollection(1)
        .Format.line.visible = msoFalse

        For rowIndex = 1 To sourceTable.Rows.count
            hexColor = sourceTable.cell(rowIndex, 1).shape.TextFrame.textRange.text
            If Len(hexColor) = 7 And left(hexColor, 1) = "#" Then
                r = CLng("&H" & Mid(hexColor, 2, 2))
                g = CLng("&H" & Mid(hexColor, 4, 2))
                b = CLng("&H" & Mid(hexColor, 6, 2))

                .Points(rowIndex).Format.Fill.ForeColor.RGB = RGB(r, g, b)
                .Points(rowIndex).Format.line.visible = msoFalse
                .Points(rowIndex).Format.line.ForeColor.RGB = RGB(255, 255, 255)
                .Points(rowIndex).Format.line.Transparency = 1
                .Points(rowIndex).Format.line.Weight = 0
            Else
                ' Skip invalid hex codes
            End If
        Next rowIndex
    End With

    ' Ta bort SOURCE-tabellen efter användning
    tableShape.Delete
    Exit Sub

ErrorHandler:
    ' Handle errors gracefully and exit
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Next
End Sub

