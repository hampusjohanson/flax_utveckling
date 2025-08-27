Attribute VB_Name = "Create_Balance_2"
Sub Create_balance_2()
    Dim pptSlide As slide
    Dim circleShape As shape
    Dim xPos As Single
    Dim yPos As Single
    Dim width As Single
    Dim height As Single

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Check if the slide is valid
    If pptSlide Is Nothing Then
        MsgBox "No active slide found.", vbExclamation
        Exit Sub
    End If

    ' Define position and size in points (cm to points conversion)
    xPos = 14.93 * 28.35  ' 14.93 cm converted to points
    yPos = 5.12 * 28.35   ' 5.12 cm converted to points
    width = 2 * 28.35      ' 2 cm converted to points (width of the circle)
    height = 2 * 28.35     ' 2 cm converted to points (height of the circle)

    ' Create the circle shape
    Set circleShape = pptSlide.Shapes.AddShape(msoShapeOval, xPos, yPos, width, height)

    ' Set name for the circle
    circleShape.Name = "BalanceCircle"

    ' Debugging: Confirm circle creation
    Debug.Print "Circle created at position (" & xPos & ", " & yPos & ") with size (" & width & ", " & height & ")"
End Sub

