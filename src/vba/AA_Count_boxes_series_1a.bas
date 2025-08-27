Attribute VB_Name = "AA_Count_boxes_series_1a"
Sub AA_rightest_leftie()
    Dim pptSlide As slide
    Dim shape As shape
    Dim targetColor As Long
    Dim leftmostBox As shape
    Dim rightmostBox As shape
    Dim minLeft As Single
    Dim maxRight As Single

    ' === Set the active slide ===
    Set pptSlide = ActiveWindow.View.slide

    ' === Define the target fill color (#88FFC2) ===
    targetColor = RGB(136, 255, 194)

    ' === Initialize variables ===
    minLeft = 999999 ' Start with a very high value
    maxRight = -1    ' Start with a very low value
    Set leftmostBox = Nothing
    Set rightmostBox = Nothing

    ' === Loop through all shapes on the slide ===
    For Each shape In pptSlide.Shapes
        ' Check if the shape has a fill and matches the target color
        If shape.Fill.visible = msoTrue Then
            If shape.Fill.ForeColor.RGB = targetColor Then
                ' Check if this is the leftmost box
                If shape.left < minLeft Then
                    minLeft = shape.left
                    Set leftmostBox = shape
                End If
                ' Check if this is the rightmost box
                If (shape.left + shape.width) > maxRight Then
                    maxRight = shape.left + shape.width
                    Set rightmostBox = shape
                End If
            End If
        End If
    Next shape

    ' === Rename the leftmost and rightmost boxes ===
    If Not leftmostBox Is Nothing Then
        leftmostBox.Name = "Leftie"
        Debug.Print "Leftmost box renamed to 'Leftie'."
    Else
        Debug.Print "No leftmost box found with the target color."
    End If

    If Not rightmostBox Is Nothing Then
        rightmostBox.Name = "Rightie"
        Debug.Print "Rightmost box renamed to 'Rightie'."
    Else
        Debug.Print "No rightmost box found with the target color."
    End If

    
End Sub

