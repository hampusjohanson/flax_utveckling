Attribute VB_Name = "Create_balance_9"
Sub Create_balance_9()
    Dim pptSlide As slide
    Dim textBoxShape As shape
    Dim newWidth As Single

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Ensure the "MyTextBox" exists
    On Error Resume Next
    Set textBoxShape = pptSlide.Shapes("MyTextBox")
    On Error GoTo 0

    If textBoxShape Is Nothing Then
        MsgBox "'MyTextBox' text box not found. Please ensure it exists.", vbExclamation
        Exit Sub
    End If

    ' Convert the new width from cm to points (1 cm = 28.35 points)
    newWidth = 5.5 * 28.35  ' 5.5 cm converted to points

    ' Change the width of the text box
    textBoxShape.width = newWidth

    ' Debugging - Confirm the change in width
    Debug.Print "Width of 'MyTextBox' changed to: " & textBoxShape.width & " points"
End Sub

