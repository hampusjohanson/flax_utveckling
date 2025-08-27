Attribute VB_Name = "Clear_Right"
Sub Clear_Right()

    Dim sld As slide
    Dim shp As shape
    Dim rightmostShape As shape
    Dim maxLeft As Single
    Dim tbl As table
    Dim r As Long, c As Long
    
    ' H�mta den aktuella sliden i redigeringsl�ge:
    Set sld = ActiveWindow.View.slide
    
    ' OM du k�r i bildspelsl�ge ist�llet, anv�nd:
    ' Set sld = SlideShowWindows(1).View.Slide

    ' Leta igenom alla shapes p� sliden och hitta den som:
    ' 1) Inneh�ller en tabell
    ' 2) Har h�gst "Left"-v�rde (d.v.s. ligger l�ngst till h�ger)
    For Each shp In sld.Shapes
        If shp.HasTable Then
            If shp.left > maxLeft Then
                maxLeft = shp.left
                Set rightmostShape = shp
            End If
        End If
    Next shp

    ' Om vi inte hittade n�gon tabell:
    If rightmostShape Is Nothing Then
        MsgBox "Ingen tabell hittades p� den h�r sliden.", vbInformation
        Exit Sub
    End If
    
    ' Referera till tabellen
    Set tbl = rightmostShape.table
    
    ' Rensa inneh�llet i cellerna fr�n rad 2 till sista raden
    For r = 2 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            tbl.cell(r, c).shape.TextFrame.textRange.text = ""
        Next c
    Next r


End Sub

