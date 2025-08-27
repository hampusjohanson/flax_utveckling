Attribute VB_Name = "AA_Count_boxes_series_2"
Sub AA_Count_boxes_series_2()
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim leftieBox As shape
    Dim newBox As shape
    Dim boxesNeeded As Integer
    Dim i As Integer
    Dim shapeList As Collection
    Dim totalBoxes As Integer
    Dim leftPos As Single
    Dim topPos As Single
    Dim spacing As Single

    ' === Set active slide ===
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
        MsgBox "BOX table not found.", vbExclamation
        Exit Sub
    End If

    ' === Read "Boxes needed" value (row 4, column 2) ===
    boxesNeeded = val(tableShape.table.cell(4, 2).shape.TextFrame.textRange.text)

    ' === Exit if no boxes are needed ===
    If boxesNeeded <= 0 Then
        
        Exit Sub
    End If

    ' === Find "Leftie" shape ===
    Set leftieBox = Nothing
    For Each shape In pptSlide.Shapes
        If shape.Name = "Leftie" Then
            Set leftieBox = shape
            Exit For
        End If
    Next shape

    ' === Check if Leftie was found ===
    If leftieBox Is Nothing Then
        MsgBox "No shape named 'Leftie' found.", vbExclamation
        Exit Sub
    End If

    ' === Store all shapes in a collection ===
    Set shapeList = New Collection
    shapeList.Add leftieBox ' Add original Leftie to the list

    ' === Duplicate "Leftie" as many times as needed ===
    For i = 1 To boxesNeeded
        Set newBox = leftieBox.Duplicate.Item(1) ' Fix: Correctly reference duplicated shape
        newBox.Name = "Leftie_Copy_" & i
        shapeList.Add newBox
    Next i

    ' === Get total number of boxes ===
    totalBoxes = shapeList.count

    ' === Get positioning values ===
    leftPos = leftieBox.left
    topPos = leftieBox.Top
    spacing = leftieBox.width + 20 ' Adjust spacing as needed

    ' === Horizontally distribute the boxes ===
    For i = 1 To totalBoxes
        shapeList(i).left = leftPos + (i - 1) * spacing
        shapeList(i).Top = topPos ' Align to the same top position
    Next i

  
    ' Delete the "SOURCE" table from the slide
    tableShape.Delete
  
End Sub

