Attribute VB_Name = "SS92"
Sub RenameTextboxByPosition()
    Dim pptSlide As slide
    Dim s As shape
    Dim targetLeft As Double, targetTop As Double
    Dim tolerance As Double
    Dim renamed As Boolean
    
    ' Define target position in points (PowerPoint uses points, 1 cm = 28.35 points)
    targetLeft = 4.27 * 28.35
    targetTop = 16.13 * 28.35
    tolerance = 5 ' Allow small variation to account for rounding

    ' Get active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Iterate through shapes to find the one at the given position
    For Each s In pptSlide.Shapes
        If Abs(s.left - targetLeft) < tolerance And Abs(s.Top - targetTop) < tolerance Then
            Debug.Print "Textbox found at position: Left=" & s.left & ", Top=" & s.Top
            s.Name = "brand_x"
            Debug.Print "Textbox renamed to 'brand_x'"
            renamed = True
            Exit For
        End If
    Next s

    ' If no textbox was found
    If Not renamed Then
        MsgBox "No textbox found at the specified position.", vbExclamation
    End If
End Sub

