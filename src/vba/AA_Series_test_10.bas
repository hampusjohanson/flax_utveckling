Attribute VB_Name = "AA_Series_test_10"
Sub AA_Series_10()
    Dim pptSlide As slide
    Dim shapeObj As shape
    Dim i As Integer

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loop backwards through shapes to avoid indexing issues when deleting
    For i = pptSlide.Shapes.count To 1 Step -1
        Set shapeObj = pptSlide.Shapes(i)

        ' Check if shape name starts with "Leftie"
        If left(shapeObj.Name, 6) = "Leftie" Then
            shapeObj.Delete
        End If
    Next i

    
End Sub

