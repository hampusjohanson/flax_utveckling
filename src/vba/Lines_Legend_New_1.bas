Attribute VB_Name = "Lines_Legend_New_1"
Public brand_count As Integer

Sub Lines_Legend_New_1()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim countVisible As Integer
    Dim sIndex As Integer

    Set pptSlide = ActiveWindow.View.slide
    countVisible = 0

    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            For sIndex = 1 To chartObject.SeriesCollection.count
                Set series = chartObject.SeriesCollection(sIndex)
                If series.Format.line.visible = msoTrue Then
                    countVisible = countVisible + 1
                End If
            Next sIndex
            Exit For
        End If
    Next chartShape

    brand_count = countVisible
    Debug.Print "Visible brands: " & brand_count
End Sub

Sub Lines_Legend_New_2()
    Dim pptSlide As slide
    Dim tblShape As shape
    Dim tbl As table
    Dim i As Integer
    Dim numRows As Integer
    Dim numCols As Integer: numCols = 2
    Dim leftPos As Single: leftPos = 25.82 * 28.35
    Dim topPos As Single: topPos = 5.14 * 28.35
    Dim width As Single: width = (0.46 + 3) * 28.35
    Dim height As Single
    Dim r As Integer, c As Integer

    Set pptSlide = ActiveWindow.View.slide

    ' Bestäm antal rader baserat på brand_count
    If brand_count = 0 Then
        MsgBox "brand_count är 0 – kör Lines_Legend_New_1 först?", vbExclamation
        Exit Sub
    End If
    numRows = brand_count
    height = 0.47 * 2 * 28.35 ' startstorlek för två rader

    ' Ta bort eventuell gammal tabell
    On Error Resume Next
    pptSlide.Shapes("Brand_List_1").Delete
    On Error GoTo 0

    ' Skapa ny tabell med 2 rader (lägger till resten efteråt)
    Set tblShape = pptSlide.Shapes.AddTable(2, numCols, leftPos, topPos, width, height)
    tblShape.Name = "Brand_List_1"
    Set tbl = tblShape.table

    ' Justera kolumnbredder
    tbl.Columns(1).width = 0.46 * 28.35
    tbl.Columns(2).width = 3 * 28.35

    ' Justera höjd på de två första
    For i = 1 To 2
        tbl.Rows(i).height = 0.47 * 28.35
    Next i

    ' Lägg till rader tills vi har rätt antal
    Do While tbl.Rows.count < numRows
        tbl.Rows.Add
        tbl.Rows(tbl.Rows.count).height = 0.47 * 28.35
    Loop

    ' Rensa innehåll och sätt font
    For r = 1 To numRows
        For c = 1 To numCols
            With tbl.cell(r, c).shape.TextFrame.textRange
                .Font.Name = "Arial"
                .Font.size = 2
                .Font.Bold = msoFalse
                .text = ""
            End With
        Next c
    Next r

    Debug.Print "Brand_List_1 skapad med " & numRows & " rader"
End Sub



Sub Lines_Legend_New_3()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer
    Dim b

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_1")
    Set tbl = shapeTbl.table

    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            For Each b In Array(ppBorderTop, ppBorderBottom, ppBorderLeft, ppBorderRight)
                tbl.cell(r, c).Borders(b).visible = msoFalse
            Next b
        Next c
    Next r

    Debug.Print "Alla kantlinjer borttagna från tabellen."
End Sub

Sub Lines_Legend_New_4()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer
    Dim b

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_1")
    Set tbl = shapeTbl.table

    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            For Each b In Array(ppBorderTop, ppBorderBottom, ppBorderLeft, ppBorderRight)
                With tbl.cell(r, c).Borders(b)
                    .ForeColor.RGB = RGB(242, 242, 241)
                    .Weight = 0.75
                    .visible = msoTrue
                End With
            Next b
        Next c
    Next r

    Debug.Print "Alla kantlinjer satta till grå färg i tabellen."
End Sub

Sub Lines_Legend_New_5()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim tbl As table
    Dim visibleList As Collection
    Dim r As Integer
    Dim fillColor As Long

    Set pptSlide = ActiveWindow.View.slide

    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            Exit For
        End If
    Next chartShape

    If chartObject Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbExclamation
        Exit Sub
    End If

    Set tbl = pptSlide.Shapes("Brand_List_1").table
    Set visibleList = GetVisibleSeriesIndexes(chartObject)

    For r = 1 To visibleList.count
        fillColor = chartObject.SeriesCollection(visibleList(r)).Format.line.ForeColor.RGB

        With tbl.cell(r, 1).shape
            .Fill.visible = msoTrue
            .Fill.ForeColor.RGB = fillColor
            .Fill.Solid
            .TextFrame.textRange.text = " "
            With .TextFrame
                .MarginTop = 2
                .MarginBottom = 2
                .MarginLeft = 2
                .MarginRight = 2
            End With
        End With

        Dim borderType As PpBorderType
        For borderType = ppBorderTop To ppBorderRight
            With tbl.cell(r, 1).Borders(borderType)
                .ForeColor.RGB = RGB(&HF2, &HF2, &HF1)
                .Weight = 6
            End With
        Next borderType

        tbl.cell(r, 2).shape.Fill.visible = msoFalse
        Dim b
        For Each b In Array(ppBorderTop, ppBorderBottom, ppBorderLeft, ppBorderRight)
            tbl.cell(r, 2).Borders(b).visible = msoFalse
        Next b
    Next r

    Debug.Print "Färgade Brand_List_1 med " & visibleList.count & " synliga serier"
End Sub


Sub Lines_Legend_New_6()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_1")
    Set tbl = shapeTbl.table

    tbl.Columns(1).width = 13.36

    For r = 1 To tbl.Rows.count
        tbl.Rows(r).height = 13.36
    Next r

    Debug.Print "Kolumn 1 satt till 0,47 x 0,47 cm"
End Sub
Sub Lines_Legend_New_7()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim tbl As table
    Dim visibleList As Collection
    Dim r As Integer
    Dim nameText As String

    Set pptSlide = ActiveWindow.View.slide

    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            Exit For
        End If
    Next chartShape

    If chartObject Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbExclamation
        Exit Sub
    End If

    Set tbl = pptSlide.Shapes("Brand_List_1").table
    Set visibleList = GetVisibleSeriesIndexes(chartObject)

    For r = 1 To visibleList.count
        nameText = chartObject.SeriesCollection(visibleList(r)).Name

        With tbl.cell(r, 2).shape.TextFrame.textRange
            .text = nameText
            .Font.size = 8
            .Font.Name = "Arial"
            .Font.color.RGB = RGB(17, 21, 66) ' #111542
            .Font.Bold = msoFalse
        End With
    Next r

    Debug.Print "Namn från " & visibleList.count & " synliga serier inmatade i Brand_List_1"
End Sub


Sub Lines_Legend_New_8()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_1")
    Set tbl = shapeTbl.table

    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            With tbl.cell(r, c).shape.TextFrame
                .MarginTop = 1.42
                .MarginBottom = 1.42
            End With
        Next c
    Next r

    Debug.Print "Top/Bottom textmarginaler satta till 0,05 cm i hela tabellen."
End Sub

