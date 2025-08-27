Attribute VB_Name = "Ld_4"
Sub Mac_LD_3()
'Fyll i samma med bold
    
    Dim pptSlide As slide
    Dim rightmostTableShape As shape
    Dim rightmostTable As table
    Dim col2Values() As String
    Dim col5Values() As String
    Dim i As Integer, j As Integer
    Dim MatchFound As Boolean
    Dim rightMostPosition As Single

    ' === Hitta tabellen längst till höger på sliden ===
    Set pptSlide = ActiveWindow.View.slide
    rightMostPosition = -99999 ' Initialt lågt värde för att hitta högerpositionen
    Set rightmostTableShape = Nothing

    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            If shape.left > rightMostPosition Then
                rightMostPosition = shape.left
                Set rightmostTableShape = shape
            End If
        End If
    Next shape

    If rightmostTableShape Is Nothing Then
        MsgBox "Ingen tabell hittades på sliden.", vbExclamation
        Exit Sub
    End If

    ' === Hämta referens till tabellen längst till höger ===
    Set rightmostTable = rightmostTableShape.table

    ' === Lägg till värden från kolumn 2 och kolumn 5 i arrayer ===
    ReDim col2Values(1 To rightmostTable.Rows.count)
    ReDim col5Values(1 To rightmostTable.Rows.count)

    For i = 1 To rightmostTable.Rows.count
        col2Values(i) = Trim(rightmostTable.cell(i, 2).shape.TextFrame.textRange.text)
        col5Values(i) = Trim(rightmostTable.cell(i, 5).shape.TextFrame.textRange.text)
    Next i

    ' === Kontrollera kolumn 2 mot kolumn 5 ===
    For i = 1 To rightmostTable.Rows.count
        If Len(col2Values(i)) > 2 And i <> 1 Then ' Kontrollera att texten är längre än 2 tecken och att det inte är rad 1
            MatchFound = False
            For j = 1 To rightmostTable.Rows.count
                If col2Values(i) = col5Values(j) Then
                    MatchFound = True
                    Exit For
                End If
            Next j

            ' Fetmarkera cell i kolumn 2 och motsvarande cell i kolumn 1 om match hittades
            If MatchFound Then
                With rightmostTable.cell(i, 2).shape.TextFrame.textRange.Font
                    .Bold = msoTrue
                End With

                ' Fetmarkera cell i samma rad i kolumn 1
                With rightmostTable.cell(i, 1).shape.TextFrame.textRange.Font
                    .Bold = msoTrue
                End With
            End If
        End If
    Next i

    ' === Kontrollera kolumn 5 mot kolumn 2 ===
    For i = 1 To rightmostTable.Rows.count
        If Len(col5Values(i)) > 2 And i <> 1 Then ' Kontrollera att texten är längre än 2 tecken och att det inte är rad 1
            MatchFound = False
            For j = 1 To rightmostTable.Rows.count
                If col5Values(i) = col2Values(j) Then
                    MatchFound = True
                    Exit For
                End If
            Next j

            ' Fetmarkera cell i kolumn 5 och motsvarande cell i kolumn 4 om match hittades
            If MatchFound Then
                With rightmostTable.cell(i, 5).shape.TextFrame.textRange.Font
                    .Bold = msoTrue
                End With

                ' Fetmarkera cell i samma rad i kolumn 4
                With rightmostTable.cell(i, 4).shape.TextFrame.textRange.Font
                    .Bold = msoTrue
                End With
            End If
        End If
    Next i
End Sub

