Attribute VB_Name = "Export_1"
Sub ExportTableInfoToFile()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim tbl As table
    Dim rowIndex As Integer, colIndex As Integer
    Dim cellInfo As String
    Dim exportText As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim operatingSystem As String
    Dim userName As String

    exportText = "" ' Initialisera exporttext

    ' H�mta den aktuella sliden
    Set pptSlide = ActiveWindow.View.slide

    ' Kontrollera operativsystem och s�tt filv�g
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
        filePath = "/Users/" & userName & "/Desktop/exported_table_info.txt"
    Else
        ' F�r Windows
        filePath = "C:\\Users\\" & Environ("USERNAME") & "\\Desktop\\exported_table_info.txt"
    End If

    ' Loopa igenom alla former p� sliden f�r att hitta tabeller
    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            Set tbl = tableShape.table
            exportText = exportText & "Tabellinfo: " & vbCrLf
            exportText = exportText & "Position (Left, Top): (" & tableShape.left & ", " & tableShape.Top & ")" & vbCrLf
            exportText = exportText & "Storlek (Width, Height): (" & tableShape.width & ", " & tableShape.height & ")" & vbCrLf
            exportText = exportText & "Rotation: " & tableShape.Rotation & " grader" & vbCrLf
            exportText = exportText & "Synlighet: " & IIf(tableShape.visible = msoTrue, "Synlig", "Dold") & vbCrLf
            exportText = exportText & "Antal rader: " & tbl.Rows.count & ", Antal kolumner: " & tbl.Columns.count & vbCrLf & vbCrLf

            ' Loopa igenom varje cell f�r att samla info
            For rowIndex = 1 To tbl.Rows.count
                For colIndex = 1 To tbl.Columns.count
                    With tbl.cell(rowIndex, colIndex)
                        cellInfo = "Rad " & rowIndex & ", Kolumn " & colIndex & ": " & vbCrLf
                        cellInfo = cellInfo & "  Text: " & .shape.TextFrame.textRange.text & vbCrLf
                        cellInfo = cellInfo & "  H�jd: " & tbl.Rows(rowIndex).height & vbCrLf
                        cellInfo = cellInfo & "  Bredd: " & tbl.Columns(colIndex).width & vbCrLf
                        cellInfo = cellInfo & "  Fontstorlek: " & .shape.TextFrame.textRange.Font.size & vbCrLf
                        cellInfo = cellInfo & "  Textf�rg: RGB(" & .shape.TextFrame.textRange.Font.color.RGB Mod 256 & ", " & _
                                    (.shape.TextFrame.textRange.Font.color.RGB \ 256) Mod 256 & ", " & _
                                    (.shape.TextFrame.textRange.Font.color.RGB \ 65536) Mod 256 & ")" & vbCrLf
                        cellInfo = cellInfo & "  Fyllf�rg: RGB(" & .shape.Fill.ForeColor.RGB Mod 256 & ", " & _
                                    (.shape.Fill.ForeColor.RGB \ 256) Mod 256 & ", " & _
                                    (.shape.Fill.ForeColor.RGB \ 65536) Mod 256 & ")" & vbCrLf

                        ' L�gg till information om kantlinjer
                        Dim borderIndex As Integer
                        Dim borderInfo As String
                        borderInfo = "  Kantlinjer:" & vbCrLf
                        For borderIndex = 1 To 4 ' Top, Left, Bottom, Right
                            If .Borders(borderIndex).visible Then
                                borderInfo = borderInfo & "    " & Choose(borderIndex, "Top", "Left", "Bottom", "Right") & ": Synlig" & vbCrLf
                            Else
                                borderInfo = borderInfo & "    " & Choose(borderIndex, "Top", "Left", "Bottom", "Right") & ": Dold" & vbCrLf
                            End If
                        Next borderIndex
                        cellInfo = cellInfo & borderInfo & vbCrLf

                        exportText = exportText & cellInfo & vbCrLf
                        Debug.Print cellInfo ' Skriv ut varje cells information i Immediate-f�nstret
                    End With
                Next colIndex
            Next rowIndex

            exportText = exportText & "------------------------------------------------" & vbCrLf & vbCrLf
        End If
    Next tableShape

    ' Skriv exporttext till fil
    fileNumber = FreeFile
    Open filePath For Output As fileNumber
    Print #fileNumber, exportText
    Close fileNumber

    Debug.Print exportText ' Skriv hela exporttexten till Immediate-f�nstret

    MsgBox "Tabellinfo har exporterats till: " & filePath, vbInformation
End Sub


