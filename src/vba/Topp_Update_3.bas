Attribute VB_Name = "Topp_Update_3"
Sub Topp_Update_3()
    Dim pptSlide As slide
    Dim targetShape As shape
    Dim targetTable As table
    Dim i As Integer
    Dim cellText As String

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

    ' === If column 11 is empty, make column 10 empty for the same row ===
    For i = 2 To targetTable.Rows.count
        cellText = Trim(targetTable.cell(i, 11).shape.TextFrame.textRange.text)
        If cellText = "" Then
            targetTable.cell(i, 10).shape.TextFrame.textRange.text = ""
        End If
    Next i

End Sub
