Attribute VB_Name = "SUB_TABLE_1"
Sub SUB_TABLE_1()
    Dim slide As slide
    Dim shape As shape
    
    ' Get the active slide
    Set slide = ActiveWindow.View.slide
    
    ' Loop through all shapes on the slide
    For Each shape In slide.Shapes
        ' Check if the shape is a table
        If shape.HasTable Then
            shape.Name = "TARGET"
           
            Exit Sub
        End If
    Next shape
    
    ' If no table is found
    MsgBox "No table found on the slide.", vbExclamation
End Sub

