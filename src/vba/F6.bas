Attribute VB_Name = "F6"
Function minValue(a As Integer, b As Integer) As Integer
    If a < b Then
        minValue = a
    Else
        minValue = b
    End If
End Function

Sub Hitta_Font()
    Dim sld As slide
    Dim shp As shape
    Dim fontUsage As Collection
    Dim slideFonts As Collection
    Dim fontName As String
    Dim i As Integer, j As Integer
    Dim tblSlide As slide
    Dim tblShape As shape
    Dim maxSlidesPerTable As Integer
    Dim tableTop As Single, tableLeft As Single
    Dim tableHeight As Single, tableWidth As Single
    Dim rgbColor As Long

    ' InstŠllningar
    maxSlidesPerTable = 20
    tableTop = 0 ' Startposition hšgst upp
    tableLeft = 0 ' Startposition lŠngst till vŠnster
    tableHeight = 150 ' Hšjd pŒ varje tabell
    tableWidth = 350 ' Bredd pŒ varje tabell
    rgbColor = RGB(17, 21, 66) ' FŠrgen fšr text

    ' Skapa Collection fšr att lagra typsnittsanvŠndning
    Set fontUsage = New Collection

    ' Loopa genom slides och samla in typsnittsinformation
    For Each sld In ActivePresentation.Slides
        Set slideFonts = New Collection
        
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    For i = 1 To shp.TextFrame.textRange.Characters.count
                        fontName = shp.TextFrame.textRange.Characters(i).Font.Name
                        On Error Resume Next
                        slideFonts.Add fontName, fontName
                        On Error GoTo 0
                    Next i
                End If
            End If
        Next shp
        
        ' LŠgg till informationen i fontUsage
        fontUsage.Add slideFonts
    Next sld

    ' Skapa en ny slide fšr tabellerna
    Set tblSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.count + 1, ppLayoutBlank)

    ' Skapa tabeller fšr slides
    Dim currentSlideIndex As Integer
    Dim currentTableIndex As Integer
    currentSlideIndex = 1
    currentTableIndex = 0

    Do While currentSlideIndex <= fontUsage.count
        ' Justera placeringen av tabeller pŒ samma slide
        Dim currentTableTop As Single
        Dim currentTableLeft As Single
        currentTableTop = tableTop
        currentTableLeft = tableLeft + (currentTableIndex * tableWidth)
        currentTableIndex = currentTableIndex + 1

        ' LŠgg till en ny tabell
        Set tblShape = tblSlide.Shapes.AddTable(minValue(maxSlidesPerTable, fontUsage.count - currentSlideIndex + 1) + 1, 5, currentTableLeft, currentTableTop, tableWidth, tableHeight)

        ' Styla tabellen
        tblShape.Fill.visible = msoFalse ' Ingen bakgrundsfŠrg
        tblShape.table.Columns(1).width = 45

        ' LŠgg till rubriker
        tblShape.table.cell(1, 1).shape.TextFrame.textRange.text = "Slide"
        tblShape.table.cell(1, 2).shape.TextFrame.textRange.text = "1st Font"
        tblShape.table.cell(1, 3).shape.TextFrame.textRange.text = "2nd Font"
        tblShape.table.cell(1, 4).shape.TextFrame.textRange.text = "3rd Font"
        tblShape.table.cell(1, 5).shape.TextFrame.textRange.text = "4th Font"

        ' Styla rubriker och text
        For j = 1 To tblShape.table.Rows.count
            For i = 1 To tblShape.table.Columns.count
                With tblShape.table.cell(j, i).shape.TextFrame.textRange
                    .Font.size = 10
                    .Font.color.RGB = rgbColor
                End With
            Next i
        Next j

        ' Fyll tabellen med data
        Dim currentTableRow As Integer
        currentTableRow = 2
        While currentTableRow <= maxSlidesPerTable + 1 And currentSlideIndex <= fontUsage.count
            tblShape.table.cell(currentTableRow, 1).shape.TextFrame.textRange.text = currentSlideIndex

            ' HŠmta och fyll typsnitt i tabellen
            Set slideFonts = fontUsage(currentSlideIndex)
            For j = 1 To 4 ' BegrŠnsa till 4 kolumner
                If j <= slideFonts.count Then
                    tblShape.table.cell(currentTableRow, j + 1).shape.TextFrame.textRange.text = slideFonts(j)
                End If
            Next j

            currentTableRow = currentTableRow + 1
            currentSlideIndex = currentSlideIndex + 1
        Wend
    Loop
    
MsgBox "All fonts are listed on the slide in the back.", vbInformation, "Font Listing Completed"

End Sub


