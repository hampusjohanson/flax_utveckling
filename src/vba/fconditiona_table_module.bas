Attribute VB_Name = "fconditiona_table_module"
Sub NeutralTable()
    Dim shp As shape
    Dim tbl As table
    Dim r As Long, c As Long

    ' Kontrollera att en tabell �r markerad
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell f�rst.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den markerade formen �r ingen tabell.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.table

    ' Loopa igenom alla celler och nollst�ll formatering
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            With tbl.cell(r, c).shape.TextFrame.textRange
                ' Ta bort fetstil och kursiv
                .Font.Bold = msoFalse
                .Font.Italic = msoFalse
                ' S�tt textf�rgen till RGB(17,21,66)
                .Font.color.RGB = RGB(17, 21, 66)
            End With

            ' Ta bort all fill (bakgrund) i cellen
            With tbl.cell(r, c).shape.Fill
                .visible = msoFalse
            End With
        Next c
    Next r

End Sub

Sub FormatAllCellBorders()
    Dim shp As shape
    Dim tbl As table
    Dim r As Long, c As Long
    Dim currentCell As cell
    
    ' 1) Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en g�ng p� tabellen) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loop igenom alla celler och applicera kantlinje p� varje cell
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            Set currentCell = tbl.cell(r, c)
            
            ' Toppkant
            With currentCell.Borders(ppBorderTop)
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66)
                .Weight = 0.25
                .DashStyle = msoLineSolid
            End With
            ' Bottenkant
            With currentCell.Borders(ppBorderBottom)
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66)
                .Weight = 0.25
                .DashStyle = msoLineSolid
            End With
            ' V�nsterkant
            With currentCell.Borders(ppBorderLeft)
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66)
                .Weight = 0.25
                .DashStyle = msoLineSolid
            End With
            ' H�gerkant
            With currentCell.Borders(ppBorderRight)
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66)
                .Weight = 0.25
                .DashStyle = msoLineSolid
            End With
        Next c
    Next r
    
  
End Sub

Sub UnderlineFirstRowThick()
    Dim shp As shape
    Dim tbl As table
    Dim c As Long
    Dim firstCell As cell
    
    ' 1) Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en g�ng p� tabellen) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loopar igenom alla kolumner i f�rsta raden och s�tter en tjock 1 pt underlinje
    For c = 1 To tbl.Columns.count
        Set firstCell = tbl.cell(1, c)
        With firstCell.Borders(ppBorderBottom)
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66)
            .Weight = 1          ' 1 pt tjocklek
            .DashStyle = msoLineSolid
        End With
    Next c
    
End Sub


Sub UnderlineLastRowThick()
    Dim shp As shape
    Dim tbl As table
    Dim lastRow As Long
    Dim c As Long
    Dim lastCell As cell

    ' 1) Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en g�ng p� tabellen) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.table
    lastRow = tbl.Rows.count

    ' 2) Loop genom varje kolumn i sista raden och s�tt en tjock 1 pt underlinje
    For c = 1 To tbl.Columns.count
        Set lastCell = tbl.cell(lastRow, c)
        With lastCell.Borders(ppBorderBottom)
            .visible = msoTrue
            .ForeColor.RGB = RGB(17, 21, 66)
            .Weight = 1           ' 1 pt tjocklek
            .DashStyle = msoLineSolid
        End With
    Next c


End Sub

Sub ClearFormattingFirstRow()
    Dim shp As shape
    Dim tbl As table
    Dim c As Long
    Dim firstCell As cell
    
    ' 1) Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera tabellen f�rst (klicka en g�ng p� tabellen) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loopa genom alla kolumner i f�rsta raden och ta bort all formatering
    For c = 1 To tbl.Columns.count
        Set firstCell = tbl.cell(1, c)
        
        ' Ta bort fyllning
        With firstCell.shape.Fill
            .visible = msoFalse
        End With
        
        ' Ta bort alla kantlinjer
        With firstCell.Borders(ppBorderTop)
            .visible = msoFalse
        End With
        With firstCell.Borders(ppBorderBottom)
            .visible = msoFalse
        End With
        With firstCell.Borders(ppBorderLeft)
            .visible = msoFalse
        End With
        With firstCell.Borders(ppBorderRight)
            .visible = msoFalse
        End With
        
        ' �terst�ll textformat: ingen fetstil, ingen kursiv, svart f�rg
        With firstCell.shape.TextFrame.textRange.Font
            .Bold = msoFalse
            .Italic = msoFalse
            .color.RGB = RGB(0, 0, 0)
        End With
    Next c
    
End Sub

' Exempel: S�tt v�nsterkanten transparent f�r alla celler i den markerade tabellen
Sub MakeLeftBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long

    ' Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en g�ng p� tabellen (s� att den f�r gr� ram) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' Loop genom alla celler och s�tt v�nsterkant transparent
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            Set cell = tbl.cell(r, c)
            With cell.Borders(ppBorderLeft)
                .visible = msoTrue
                .Transparency = 1    ' 100% transparent
            End With
        Next c
    Next r

End Sub




' ===================================================================
' 1) G�r h�gerkanten transparent f�r alla celler i den markerade tabellen
' ===================================================================
Sub MakeRightBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long

    ' Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en g�ng p� tabellen (s� att den f�r gr� ram) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' Loop genom alla celler och s�tt h�gerkant transparent
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            Set cell = tbl.cell(r, c)
            With cell.Borders(ppBorderRight)
                .visible = msoTrue
                .Transparency = 1    ' 100% transparent
            End With
        Next c
    Next r

End Sub

Sub MakeMiddleVerticalBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long
    Dim lastCol As Long

    ' Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en g�ng p� tabellen (s� att den f�r gr� ram) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table
    lastCol = tbl.Columns.count

    ' Loop genom rader och kolumner (1 till lastCol-1) f�r att n� inre vertikala linjer
    For r = 1 To tbl.Rows.count
        For c = 1 To lastCol - 1
            Set cell = tbl.cell(r, c)
            With cell.Borders(ppBorderRight)
                .visible = msoTrue
                .Transparency = 1    ' 100% transparent
            End With
        Next c
    Next r


End Sub


Sub MakeMiddleHorizontalBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long
    Dim lastRow As Long

    ' Kontrollera att en tabell �r markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en g�ng p� tabellen (s� att den f�r gr� ram) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table
    lastRow = tbl.Rows.count

    ' Loop genom rader 1 to lastRow-1 och alla kolumner f�r att n� inre horisontella linjer
    For r = 1 To lastRow - 1
        For c = 1 To tbl.Columns.count
            Set cell = tbl.cell(r, c)
            With cell.Borders(ppBorderBottom)
                .visible = msoTrue
                .Transparency = 1    ' 100% transparent
            End With
        Next c
    Next r

End Sub

Sub RemoveAllBorders_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long

    ' 1) Kontrollera att en tabell �r markerad
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en g�ng p� tabellen (s� att den f�r gr� ram) och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen �r ingen tabell. Klicka p� en tabell och f�rs�k igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' 2) Loop �ver samtliga celler och ta bort samtliga fyra kanter
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            Set cell = tbl.cell(r, c)
            With cell.Borders(ppBorderTop)
                .visible = msoFalse
            End With
            With cell.Borders(ppBorderBottom)
                .visible = msoFalse
            End With
            With cell.Borders(ppBorderLeft)
                .visible = msoFalse
            End With
            With cell.Borders(ppBorderRight)
                .visible = msoFalse
            End With
        Next c
    Next r

    ' 3) Ta �ven bort inre �grid�-linjer genom att d�lja bottenkanter p� alla men sista raden,
    '    och h�gerkanter p� alla men sista kolumnen. Men oftast r�cker steg 2).
    '    (Obs: detta �r egentligen redundant om man redan har dolt alla fyra kanter i varje cell)


End Sub



