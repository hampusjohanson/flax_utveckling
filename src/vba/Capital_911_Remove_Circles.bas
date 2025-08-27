Attribute VB_Name = "Capital_911_Remove_Circles"
Sub Mac_Cap_Remove_Circles()
    Dim pptSlide As slide
    Dim shapeToDelete As shape
    Dim shapeIndex As Integer
    Dim deletedCount As Integer

    ' HŠmta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loopa baklŠnges genom alla former pŒ sliden och ta bort cirklar med namn "CircleX"
    deletedCount = 0
    For shapeIndex = pptSlide.Shapes.count To 1 Step -1
        Set shapeToDelete = pptSlide.Shapes(shapeIndex)
        If left(shapeToDelete.Name, 6) = "Circle" Then
            shapeToDelete.Delete
            deletedCount = deletedCount + 1
        End If
    Next shapeIndex

    ' Meddelande efter slutfšrande
    If deletedCount > 0 Then
       Else
        MsgBox "Inga cirklar med namn 'CircleX' hittades.", vbExclamation
    End If
End Sub

