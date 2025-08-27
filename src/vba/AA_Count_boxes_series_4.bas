Attribute VB_Name = "AA_Count_boxes_series_4"
Sub AA_Count_boxes_series_4()
    Dim pptSlide As slide
    Dim shape As shape
    Dim tableShape As shape

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Find the "BOX" table ===
    Set tableShape = Nothing
    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            If shape.table.cell(1, 1).shape.TextFrame.textRange.text = "Metric" Then
                Set tableShape = shape
                Exit For
            End If
        End If
    Next shape

    ' === Check if table was found ===
    If tableShape Is Nothing Then
       
        Exit Sub
    End If

    ' === Delete the table ===
    tableShape.Delete

  
End Sub

