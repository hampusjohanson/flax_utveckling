Attribute VB_Name = "Clear_Left"
Sub Clear_Left()

    Dim sld As slide
    Dim shp As shape
    Dim leftmostShape As shape
    Dim minLeft As Single
    Dim tbl As table
    Dim r As Long, c As Long
    
    ' S�tt minLeft till ett stort v�rde s� att f�rsta tabell vi hittar automatiskt blir minst
    minLeft = 999999
    
    ' H�mta den aktuella sliden i redigeringsl�ge:
    Set sld = ActiveWindow.View.slide
    
    ' OM du k�r i bildspelsl�ge ist�llet, anv�nd:
    ' Set sld = SlideShowWindows(1).View.Slide

    ' Leta igenom alla shapes p� sliden och hitta den som:
    ' 1) Inneh�ller en tabell
    ' 2) Har l�gst "Left"-v�rde (d.v.s. ligger l�ngst till v�nster)
    For Each shp In sld.Shapes
        If shp.HasTable Then
            If shp.left < minLeft Then
                minLeft = shp.left
                Set leftmostShape = shp
            End If
        End If
    Next shp

    ' Om vi inte hittade n�gon tabell:
    If leftmostShape Is Nothing Then
        MsgBox "Ingen tabell hittades p� den h�r sliden.", vbInformation
        Exit Sub
    End If
    
    ' Referera till tabellen
    Set tbl = leftmostShape.table
    
    ' Rensa inneh�llet i cellerna fr�n rad 2 till sista raden
    For r = 2 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            tbl.cell(r, c).shape.TextFrame.textRange.text = ""
        Next c
    Next r


End Sub

