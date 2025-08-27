Attribute VB_Name = "AA_Delete_Brands"
Sub DeleteBrandsTable()
    Dim pptSlide As slide
    Dim tableShape As shape

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the table named "Brands"
    For Each tableShape In pptSlide.Shapes
        If tableShape.Name = "Brands" Then
            ' Delete the table
            tableShape.Delete
            
            Exit Sub
        End If
    Next tableShape

    ' If the table was not found
   End Sub

