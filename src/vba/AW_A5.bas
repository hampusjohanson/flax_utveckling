Attribute VB_Name = "AW_A5"
Sub DeleteChartsWithConfirmation3()
    Dim sld As slide
    Dim shp As shape
    Dim i As Integer
    Dim response As VbMsgBoxResult

    On Error Resume Next
    Set sld = ActiveWindow.View.slide
    On Error GoTo 0

    If sld Is Nothing Then Exit Sub

    For i = sld.Shapes.count To 1 Step -1
        Set shp = sld.Shapes(i)

        If shp.Type = msoChart Then
            response = MsgBox("Do you want to delete this chart?", vbYesNo + vbQuestion, "Delete Chart?")
            If response = vbYes Then shp.Delete
        End If
    Next i
End Sub


