Attribute VB_Name = "Lines_Legend_New_2"
Public brand_count As Integer

Sub Lines_Legend_New_B1()
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

Sub Lines_Legend_New_B2()
    Dim pptSlide As slide
    Dim tblShape As shape
    Dim tbl As table
    Dim i As Integer
    Dim initialRows As Integer: initialRows = 2
    Dim finalRows As Integer: finalRows = 6
    Dim numCols As Integer: numCols = 2
    Dim leftPos As Single: leftPos = 29.68 * 28.35
    Dim topPos As Single: topPos = 5.14 * 28.35
    Dim width As Single: width = (0.46 + 3) * 28.35
    Dim height As Single: height = 0.47 * initialRows * 28.35

    Set pptSlide = ActiveWindow.View.slide
    On Error Resume Next
    pptSlide.Shapes("Brand_List_2").Delete
    On Error GoTo 0

    Set tblShape = pptSlide.Shapes.AddTable(initialRows, numCols, leftPos, topPos, width, height)
    tblShape.Name = "Brand_List_2"
    Set tbl = tblShape.table

    tbl.Columns(1).width = 0.46 * 28.35
    tbl.Columns(2).width = 3 * 28.35

    For i = 1 To initialRows
        tbl.Rows(i).height = 0.47 * 28.35
    Next i
    For i = initialRows + 1 To finalRows
        tbl.Rows.Add
        tbl.Rows(i).height = 0.47 * 28.35
    Next i

    Dim r As Integer, c As Integer
    For r = 1 To finalRows
        For c = 1 To numCols
            With tbl.cell(r, c).shape.TextFrame.textRange
                .Font.Name = "Arial"
                .Font.size = 2
                .Font.Bold = msoFalse
                .text = ""
            End With
        Next c
    Next r

    Debug.Print "Brand_List_2 created with fixed size and 6 rows"
End Sub

Sub Lines_Legend_New_B3()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer
    Dim b

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_2")
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

Sub Lines_Legend_New_B4()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer
    Dim b

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_2")
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

Sub Lines_Legend_New_B5()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim tbl As table
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

    Set tbl = pptSlide.Shapes("Brand_List_2").table

    For r = 1 To 6
        If (r + 6) <= chartObject.SeriesCollection.count Then
            fillColor = chartObject.SeriesCollection(r + 6).Format.line.ForeColor.RGB
        Else
            fillColor = RGB(200, 200, 200)
        End If

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

    Debug.Print "Brand_List_2 färdig: rätt färger från serie 7–12"
End Sub




Sub Lines_Legend_New_B6()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_2")
    Set tbl = shapeTbl.table

    tbl.Columns(1).width = 13.36

    For r = 1 To tbl.Rows.count
        tbl.Rows(r).height = 13.36
    Next r

    Debug.Print "Kolumn 1 satt till 0,47 x 0,47 cm"
End Sub

Sub Lines_Legend_New_B7()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim tbl As table
    Dim r As Integer

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

    Set tbl = pptSlide.Shapes("Brand_List_2").table

    For r = 1 To 6
        Dim nameText As String
        If (r + 6) <= chartObject.SeriesCollection.count Then
            nameText = chartObject.SeriesCollection(r + 6).Name
        Else
            nameText = "Varumärke " & (r + 6)
        End If

        With tbl.cell(r, 2).shape.TextFrame.textRange
            .text = nameText
            .Font.size = 8
            .Font.Name = "Arial"
            .Font.color.RGB = RGB(17, 21, 66) ' #111542
            .Font.Bold = msoFalse
        End With
    Next r

    Debug.Print "Varumärkesnamn (serie 7–12) inmatade i Brand_List_2 kolumn 2"
End Sub


Sub Lines_Legend_New_B8()
    Dim pptSlide As slide
    Dim shapeTbl As shape
    Dim tbl As table
    Dim r As Integer, c As Integer

    Set pptSlide = ActiveWindow.View.slide
    Set shapeTbl = pptSlide.Shapes("Brand_List_2")
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


