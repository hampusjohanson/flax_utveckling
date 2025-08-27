Attribute VB_Name = "AA_Count_boxes_series_3"
Sub AA_Count_boxes_series_3()
    Dim pptSlide As slide
    Dim leftieBox As shape
    Dim rightieBox As shape
    Dim shape As shape
    Dim shapeList As Collection
    Dim totalBoxes As Integer
    Dim leftPos As Single
    Dim rightPos As Single
    Dim spacing As Single
    Dim i As Integer

    ' === Set active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Initialize collection ===
    Set shapeList = New Collection

    ' === Find "Leftie" and "Rightie" shapes ===
    Set leftieBox = Nothing
    Set rightieBox = Nothing

    For Each shape In pptSlide.Shapes
        If shape.Name = "Leftie" Then
            Set leftieBox = shape
        ElseIf shape.Name = "Rightie" Then
            Set rightieBox = shape
        End If
    Next shape

    ' === Check if both Leftie and Rightie were found ===
    If leftieBox Is Nothing Or rightieBox Is Nothing Then
        MsgBox "Either 'Leftie_1' or 'Rightie' was not found.", vbExclamation
        Exit Sub
    End If

    ' === Debug: Print found positions ===
    Debug.Print "Leftie Position: " & leftieBox.left
    Debug.Print "Rightie Position: " & rightieBox.left

    ' === Ensure Leftie is actually to the left of Rightie ===
    If leftieBox.left > rightieBox.left Then
        MsgBox "'Leftie_1' is positioned to the right of 'Rightie'. Please correct their positions.", vbExclamation
        Exit Sub
    End If

    ' === Add all shapes between Leftie and Rightie to the collection (including them) ===
    For Each shape In pptSlide.Shapes
        If (shape.left >= leftieBox.left And shape.left <= rightieBox.left + rightieBox.width) Then
            shapeList.Add shape
        End If
    Next shape

    ' === Ensure at least two shapes are found ===
    totalBoxes = shapeList.count
    Debug.Print "Total boxes selected: " & totalBoxes

    If totalBoxes < 2 Then
        MsgBox "Not enough shapes found between 'Leftie_1' and 'Rightie' to distribute.", vbExclamation
        Exit Sub
    End If

    ' === Get positioning values ===
    leftPos = leftieBox.left
    rightPos = rightieBox.left
    spacing = (rightPos - leftPos) / (totalBoxes - 1) ' Calculate spacing

    ' === Rename and horizontally distribute the boxes ===
    For i = 1 To totalBoxes
        shapeList(i).left = leftPos + (i - 1) * spacing
        shapeList(i).Name = "Leftie_" & i
        Debug.Print "Box " & i & " renamed to: " & shapeList(i).Name & " positioned at: " & shapeList(i).left
    Next i

   
End Sub

