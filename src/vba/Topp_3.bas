Attribute VB_Name = "Topp_3"
Sub Topp_3()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim rowIndex As Integer
    Dim cellValue As String

    ' Hämta den aktiva sliden
    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    If pptSlide Is Nothing Then
        MsgBox "Ingen aktiv slide hittades. Kontrollera att du har en presentation öppen.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Hämta tabellen "TARGET"
    On Error Resume Next
    Set targetShape = pptSlide.Shapes("TARGET")
    If targetShape Is Nothing Or Not targetShape.HasTable Then
        MsgBox "Kunde inte hitta tabellen 'TARGET'. Kontrollera att tabellen finns på sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set targetTable = targetShape.table

    ' Loopa igenom alla rader och justera borders
    For rowIndex = 1 To targetTable.Rows.count
        cellValue = Trim(targetTable.cell(rowIndex, 11).shape.TextFrame.textRange.text)

        ' Om kolumn 11 är tom
        If cellValue = "" Then
            ' Ta bort alla borders för kolumn 10 och 11 utom topborders
            Dim colIndex As Integer
            For colIndex = 10 To 11
                With targetTable.cell(rowIndex, colIndex)
                    .Borders(ppBorderLeft).visible = msoFalse
                    .Borders(ppBorderRight).visible = msoFalse
                    .Borders(ppBorderBottom).visible = msoFalse
                    .Borders(ppBorderTop).visible = msoTrue ' Behåll toppborders

                    ' Extra säkerhet - nollställ linjetjocklek
                    .Borders(ppBorderLeft).Weight = 0
                    .Borders(ppBorderRight).Weight = 0
                    .Borders(ppBorderBottom).Weight = 0
                End With
            Next colIndex
        End If
    Next rowIndex


End Sub



