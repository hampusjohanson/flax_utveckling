Attribute VB_Name = "Create_balance_6"
Sub Create_balance_6()
    Dim pptSlide As slide
    Dim balanceCircle As shape

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Ensure that the "BalanceCircle" exists
    On Error Resume Next
    Set balanceCircle = pptSlide.Shapes("BalanceCircle")
    On Error GoTo 0

    If balanceCircle Is Nothing Then
        MsgBox "The 'BalanceCircle' was not found. Please ensure it exists on the slide.", vbExclamation
        Exit Sub
    End If

    ' Set the font color inside the circle to RGB(17, 21, 66) (Dark blue color)
    balanceCircle.TextFrame.textRange.Font.color = RGB(17, 21, 66)

    ' Debugging - Verify the text color inside the circle
    Debug.Print "Font color inside 'BalanceCircle' set to: RGB(17, 21, 66)"
End Sub

