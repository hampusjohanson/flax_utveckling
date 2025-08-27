Attribute VB_Name = "Lines_Legend_Delete"
Sub Lines_Legend_Delete_Tablesa()
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide

    On Error Resume Next
    pptSlide.Shapes("Brand_List_1").Delete
    pptSlide.Shapes("Brand_List_2").Delete
    On Error GoTo 0

    Debug.Print "Eventuella Brand_List-tabeller borttagna från sliden."
End Sub


Sub Lines_Legend_Remove_Last_Row()
    Dim pptSlide As slide
    Dim tbl As table
    Dim shapeTbl As shape
    Dim lastRow As Integer

    Set pptSlide = ActiveWindow.View.slide

    On Error GoTo HandleError
    Set shapeTbl = pptSlide.Shapes("Brand_List_2")
    Set tbl = shapeTbl.table

    lastRow = tbl.Rows.count
    If lastRow > 0 Then
        tbl.Rows(lastRow).Delete
        Debug.Print "Sista raden i Brand_List_2 togs bort."
    Else
        MsgBox "Tabellen har inga rader att ta bort.", vbExclamation
    End If
    Exit Sub

HandleError:
    MsgBox "Kunde inte hitta eller modifiera Brand_List_2.", vbCritical
End Sub

