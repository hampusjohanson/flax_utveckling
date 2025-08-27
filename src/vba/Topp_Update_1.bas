Attribute VB_Name = "Topp_Update_1"
Sub Topp_Update_1()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim i As Integer, j As Integer

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Find the TARGET table ===
    On Error Resume Next
    Set targetShape = pptSlide.Shapes("TARGET")
    If targetShape Is Nothing Or Not targetShape.HasTable Then
        MsgBox "No table found named TARGET.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set targetTable = targetShape.table

    ' === Remove borders from column 1-8, row 11 and downward ===
    For i = 12 To targetTable.Rows.count
        For j = 1 To 8
            With targetTable.cell(i, j).Borders
                .Item(ppBorderRight).visible = msoFalse
                .Item(ppBorderBottom).visible = msoFalse
                .Item(ppBorderLeft).visible = msoFalse
                ' Ensure line weight is zero for extra security
                .Item(ppBorderRight).Weight = 0
                .Item(ppBorderBottom).Weight = 0
                .Item(ppBorderLeft).Weight = 0
            End With
        Next j
    Next i

End Sub

