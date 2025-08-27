Attribute VB_Name = "Lines_18"
Sub Lines_18()
    Dim pptSlide As slide
    Dim legendTable As shape
    Dim brandName As String
    Dim brandCount As Integer
    Dim requiredRow As Integer: requiredRow = 6
    Dim requiredCol As Integer: requiredCol = 2
    Dim currentRows As Integer
    Dim i As Integer

    ' Din h�rdkodade input
    brandCount = 12
    brandName = "Varum�rke 6"

    Set pptSlide = ActiveWindow.View.slide

    If brandCount <> 12 Then Exit Sub

    ' F�rs�k hitta tabellen
    On Error Resume Next
    Set legendTable = pptSlide.Shapes("Brand_List_1")
    On Error GoTo 0

    If legendTable Is Nothing Or Not legendTable.HasTable Then
        MsgBox "Tabellen 'Brand_List_1' finns inte eller �r inte en tabell.", vbExclamation
        Exit Sub
    End If

    currentRows = legendTable.table.Rows.count

    ' L�gg till rader om det beh�vs
    If currentRows < requiredRow Then
        For i = currentRows + 1 To requiredRow
            legendTable.table.Rows.Add
        Next i
    End If

    ' Kontrollera �ven kolumner
    If legendTable.table.Columns.count < requiredCol Then
        MsgBox "'Brand_List_1' har f�r f� kolumner.", vbExclamation
        Exit Sub
    End If

    ' S�tt text i cell (6, 2)
    With legendTable.table.cell(requiredRow, requiredCol).shape.TextFrame.textRange
        .text = brandName
        .Font.size = 8
        .Font.Name = "Arial"
        .Font.Bold = msoFalse
    End With
End Sub

