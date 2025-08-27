Attribute VB_Name = "Topp_Update_2"

Sub Topp_Update_2()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim i As Integer

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

    ' === Populate column 10 with sequential numbers starting from 31. ===
    For i = 2 To targetTable.Rows.count
        targetTable.cell(i, 10).shape.TextFrame.textRange.text = CStr(29 + i) & "."
    Next i

  
End Sub

