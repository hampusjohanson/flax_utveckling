Attribute VB_Name = "LD_7"
Sub Mac_LD_6()
    Dim pptSlide As slide
    Dim rightmostTableShape As shape
    Dim rightmostTable As table
    Dim i As Integer, j As Integer
    Dim swedishCount As Integer, englishCount As Integer
    Dim slideText As String
    Dim rightMostPosition As Single
    Dim boldType1 As Boolean, boldType2 As Boolean
    Dim isBoldCol2 As Boolean, isBoldCol5 As Boolean
    Dim detectedLanguage As String
    Dim tableWidth As Single, tableLeft As Single, tableTop As Single
    Dim textBox As shape
    Dim foundBoldText As String
    Dim boldFound As Boolean

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

    If rightmostTableShape Is Nothing Then Exit Sub ' Ingen tabell hittades

    ' === Hämta referens till tabellen längst till höger ===
    Set rightmostTable = rightmostTableShape.table
    tableWidth = rightmostTableShape.width
    tableLeft = rightmostTableShape.left
    tableTop = rightmostTableShape.Top + rightmostTableShape.height ' Topplacering för textrutan

    ' === Kontrollera språk i slide-text ===
    swedishCount = 0
    englishCount = 0
    slideText = ""
    For Each shape In pptSlide.Shapes
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                slideText = slideText & " " & shape.TextFrame.textRange.text
            End If
        End If
    Next shape
    slideText = LCase(slideText)

    ' Kontrollera svenska
    If InStr(slideText, "œ") > 0 Or InStr(slideText, "š") > 0 Or InStr(slideText, "š") > 0 Then swedishCount = swedishCount + 1
    If InStr(slideText, "dolda") > 0 Then swedishCount = swedishCount + 1
    If InStr(slideText, "och") > 0 Then swedishCount = swedishCount + 1

    ' Kontrollera engelska
    If InStr(slideText, "key") > 0 Then englishCount = englishCount + 1
    If InStr(slideText, "stated") > 0 Then englishCount = englishCount + 1
    If InStr(slideText, "true") > 0 Then englishCount = englishCount + 1

    If swedishCount > englishCount Then
        detectedLanguage = "Swedish"
    Else
        detectedLanguage = "English"
    End If

    ' === Identifiera Typ 1 eller Typ 2 ===
    boldType1 = False
    boldType2 = False
    foundBoldText = ""
    boldFound = False

    ' Hitta första fetmarkerade cellen i kolumn 2
    For i = 1 To rightmostTable.Rows.count
        If rightmostTable.cell(i, 2).shape.TextFrame.textRange.Font.Bold Then
            foundBoldText = Trim(rightmostTable.cell(i, 2).shape.TextFrame.textRange.text)
            boldFound = True
            Exit For
        End If
    Next i

    If boldFound Then
        ' Sök i kolumn 5 efter samma textsträng
        For j = 1 To rightmostTable.Rows.count
            If Trim(rightmostTable.cell(j, 5).shape.TextFrame.textRange.text) = foundBoldText Then
                boldType1 = True ' Typ 1 hittad
                Exit For
            End If
        Next j
    End If

    ' Om ingen match hittades eller inga fetmarkerade celler fanns, sätt till Typ 2
    If Not boldType1 Then boldType2 = True

    ' Om ingen fetmarkerad cell hittades, visa meddelande och avsluta
    If Not boldFound Then
        MsgBox "Nothing is bold here", vbExclamation
        Exit Sub
    End If

    ' === Lägg till textruta med anpassad text ===
    Set textBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, tableLeft, tableTop + 5.67, tableWidth, 28.35) ' 0,2 cm höjd = 5.67 punkter
    textBox.Name = "topp_text_ruta" ' Ge textrutan ett namn
    textBox.TextFrame.textRange.text = "" ' Placeholder

    If detectedLanguage = "Swedish" And boldType1 Then
        textBox.TextFrame.textRange.text = "Fetmarkering: Topp 10 på båda listor."
    ElseIf detectedLanguage = "Swedish" And boldType2 Then
        textBox.TextFrame.textRange.text = "Fetmarkering: Endast topp 10 på den ena listan."
    ElseIf detectedLanguage = "English" And boldType1 Then
        textBox.TextFrame.textRange.text = "Bold: Top 10 in both lists."
    ElseIf detectedLanguage = "English" And boldType2 Then
        textBox.TextFrame.textRange.text = "Bold: Top 10 in only one list."
    End If

    ' Gör "Fetmarkering:" eller "Bold:" fetmarkerat
    With textBox.TextFrame.textRange
        If detectedLanguage = "Swedish" Then
            .Characters(1, 13).Font.Bold = msoTrue
        Else
            .Characters(1, 5).Font.Bold = msoTrue
        End If

        ' Ställ in fontstorlek och färg
        .Font.size = 10
        .Font.color = RGB(17, 21, 66)
    End With

    ' Justera textrutans egenskaper
    With textBox
        .TextFrame.textRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextFrame.MarginLeft = 5
        .TextFrame.MarginRight = 5
    End With
End Sub

