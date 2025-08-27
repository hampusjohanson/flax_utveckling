Attribute VB_Name = "fconditiona_table_module"
Sub NeutralTable()
    Dim shp As shape
    Dim tbl As table
    Dim r As Long, c As Long

    ' Kontrollera att en tabell är markerad
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell först.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den markerade formen är ingen tabell.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.table

    ' Loopa igenom alla celler och nollställ formatering
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            With tbl.cell(r, c).shape.TextFrame.textRange
                ' Ta bort fetstil och kursiv
                .Font.Bold = msoFalse
                .Font.Italic = msoFalse
                ' Sätt textfärgen till RGB(17,21,66)
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
    
    ' 1) Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en gång på tabellen) och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loop igenom alla celler och applicera kantlinje på varje cell
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
            ' Vänsterkant
            With currentCell.Borders(ppBorderLeft)
                .visible = msoTrue
                .ForeColor.RGB = RGB(17, 21, 66)
                .Weight = 0.25
                .DashStyle = msoLineSolid
            End With
            ' Högerkant
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
    
    ' 1) Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en gång på tabellen) och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loopar igenom alla kolumner i första raden och sätter en tjock 1 pt underlinje
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

    ' 1) Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera en tabell (klicka en gång på tabellen) och försök igen.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.table
    lastRow = tbl.Rows.count

    ' 2) Loop genom varje kolumn i sista raden och sätt en tjock 1 pt underlinje
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
    
    ' 1) Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera tabellen först (klicka en gång på tabellen) och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    
    Set tbl = shp.table
    
    ' 2) Loopa genom alla kolumner i första raden och ta bort all formatering
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
        
        ' Återställ textformat: ingen fetstil, ingen kursiv, svart färg
        With firstCell.shape.TextFrame.textRange.Font
            .Bold = msoFalse
            .Italic = msoFalse
            .color.RGB = RGB(0, 0, 0)
        End With
    Next c
    
End Sub

' Exempel: Sätt vänsterkanten transparent för alla celler i den markerade tabellen
Sub MakeLeftBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long

    ' Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en gång på tabellen (så att den får grå ram) och försök igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' Loop genom alla celler och sätt vänsterkant transparent
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
' 1) Gör högerkanten transparent för alla celler i den markerade tabellen
' ===================================================================
Sub MakeRightBordersTransparent_SelectedTable()
    Dim shp As shape
    Dim tbl As table
    Dim cell As cell
    Dim r As Long, c As Long

    ' Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en gång på tabellen (så att den får grå ram) och försök igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' Loop genom alla celler och sätt högerkant transparent
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

    ' Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en gång på tabellen (så att den får grå ram) och försök igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table
    lastCol = tbl.Columns.count

    ' Loop genom rader och kolumner (1 till lastCol-1) för att nå inre vertikala linjer
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

    ' Kontrollera att en tabell är markerad som Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en gång på tabellen (så att den får grå ram) och försök igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table
    lastRow = tbl.Rows.count

    ' Loop genom rader 1 to lastRow-1 och alla kolumner för att nå inre horisontella linjer
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

    ' 1) Kontrollera att en tabell är markerad
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Klicka en gång på tabellen (så att den får grå ram) och försök igen.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not shp.HasTable Then
        MsgBox "Den valda formen är ingen tabell. Klicka på en tabell och försök igen.", vbExclamation
        Exit Sub
    End If
    Set tbl = shp.table

    ' 2) Loop över samtliga celler och ta bort samtliga fyra kanter
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

    ' 3) Ta även bort inre “grid”-linjer genom att dölja bottenkanter på alla men sista raden,
    '    och högerkanter på alla men sista kolumnen. Men oftast räcker steg 2).
    '    (Obs: detta är egentligen redundant om man redan har dolt alla fyra kanter i varje cell)


End Sub



