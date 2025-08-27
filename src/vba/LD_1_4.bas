Attribute VB_Name = "LD_1_4"
Sub DeleteChartsWithConfirmation()
    Dim sld As slide
    Dim shp As shape
    Dim i As Integer
    Dim response As VbMsgBoxResult
    
    Set sld = ActiveWindow.View.slide

    ' Vi loopar baklänges eftersom shapes tas bort under loopen
    For i = sld.Shapes.count To 1 Step -1
        Set shp = sld.Shapes(i)
        
        If shp.Type = msoChart Then
            ' Markera diagrammet
            shp.Select
            
            response = MsgBox("Do you want to delete this chart?", vbYesNo + vbQuestion, "Delete Chart?")
            
            If response = vbYes Then
                shp.Delete
            End If
        End If
    Next i
End Sub

