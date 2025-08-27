Attribute VB_Name = "Clear_Right"
Sub Clear_Right()

    Dim sld As slide
    Dim shp As shape
    Dim rightmostShape As shape
    Dim maxLeft As Single
    Dim tbl As table
    Dim r As Long, c As Long
    
    ' Hämta den aktuella sliden i redigeringsläge:
    Set sld = ActiveWindow.View.slide
    
    ' OM du kör i bildspelsläge istället, använd:
    ' Set sld = SlideShowWindows(1).View.Slide

    ' Leta igenom alla shapes på sliden och hitta den som:
    ' 1) Innehåller en tabell
    ' 2) Har högst "Left"-värde (d.v.s. ligger längst till höger)
    For Each shp In sld.Shapes
        If shp.HasTable Then
            If shp.left > maxLeft Then
                maxLeft = shp.left
                Set rightmostShape = shp
            End If
        End If
    Next shp

    ' Om vi inte hittade någon tabell:
    If rightmostShape Is Nothing Then
        MsgBox "Ingen tabell hittades på den här sliden.", vbInformation
        Exit Sub
    End If
    
    ' Referera till tabellen
    Set tbl = rightmostShape.table
    
    ' Rensa innehållet i cellerna från rad 2 till sista raden
    For r = 2 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            tbl.cell(r, c).shape.TextFrame.textRange.text = ""
        Next c
    Next r


End Sub

