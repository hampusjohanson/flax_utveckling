Attribute VB_Name = "ToppA_0"
Sub ToppA_0()
    Dim currentSlide As slide
    Dim shape As shape
    Dim tableNamesToDelete As Variant
    Dim tableName As Variant

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' List of table names to delete
    tableNamesToDelete = Array("long_stronger", "long_weaker", "border_table_weaker", "border_table")

    ' Loop through all shapes on the slide and delete matching tables
    For Each shape In currentSlide.Shapes
        For Each tableName In tableNamesToDelete
            If shape.Name = tableName Then
                shape.Delete
                Exit For
            End If
        Next tableName
    Next shape

End Sub

Sub ToppA_01()
    Dim currentSlide As slide
    Dim shape As shape
    Dim tableNamesToDelete As Variant
    Dim tableName As Variant

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' List of table names to delete
    tableNamesToDelete = Array("long_weaker")

    ' Loop through all shapes on the slide and delete matching tables
    For Each shape In currentSlide.Shapes
        For Each tableName In tableNamesToDelete
            If shape.Name = tableName Then
                shape.Delete
                Exit For
            End If
        Next tableName
    Next shape

End Sub

Sub ToppA_02()
    Dim currentSlide As slide
    Dim shape As shape
    Dim tableNamesToDelete As Variant
    Dim tableName As Variant

    ' Get the active slide
    Set currentSlide = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)

    ' List of table names to delete
    tableNamesToDelete = Array("border_table_weaker")

    ' Loop through all shapes on the slide and delete matching tables
    For Each shape In currentSlide.Shapes
        For Each tableName In tableNamesToDelete
            If shape.Name = tableName Then
                shape.Delete
                Exit For
            End If
        Next tableName
    Next shape

End Sub

