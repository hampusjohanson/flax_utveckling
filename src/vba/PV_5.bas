Attribute VB_Name = "PV_5"
Sub PV_5()
    ' Variabler
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim sourceTable As table
    Dim rowIndex As Integer
    Dim lastColumn As Integer

    ' === Hitta tabellen längst till höger på sliden ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing
    lastColumn = 0

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            If tableShape.left > lastColumn Then
                lastColumn = tableShape.left
                Set sourceTable = tableShape.table
            End If
        End If
    Next tableShape

    If sourceTable Is Nothing Then
        MsgBox "Ingen tabell hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' === Rensa kolumn 2 och 5 från rad 2 och nedåt ===
    For rowIndex = 2 To sourceTable.Rows.count
        sourceTable.cell(rowIndex, 2).shape.TextFrame.textRange.text = ""
        sourceTable.cell(rowIndex, 5).shape.TextFrame.textRange.text = ""
    Next rowIndex

End Sub

