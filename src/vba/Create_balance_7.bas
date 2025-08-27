Attribute VB_Name = "Create_balance_7"
Sub Create_balance_7()
    Dim pptSlide As slide
    Dim textBoxShape As shape
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

    ' Convert the position and size from cm to points (1 cm = 28.35 points)
    xPos = 17.27 * 28.35  ' 17.27 cm converted to points
    yPos = 5.57 * 28.35   ' 5.57 cm converted to points
    width = 5.5 * 28.35    ' 5.5 cm converted to points (width of the text box)
    height = 1.5 * 28.35   ' 1.5 cm converted to points (height of the text box)

    ' Create a text box at the calculated position and size
    On Error Resume Next
    Set textBoxShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, width, height) ' Position based on cm-to-point conversion
    On Error GoTo 0

    ' Check if the text box was created
    If textBoxShape Is Nothing Then
        MsgBox "Failed to create the text box. Please try again.", vbExclamation
        Exit Sub
    End If

    ' Set the name of the text box
    textBoxShape.Name = "MyTextBox"

End Sub

