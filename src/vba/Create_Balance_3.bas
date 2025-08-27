Attribute VB_Name = "Create_Balance_3"
Sub Create_balance_3()
    Dim pptSlide As slide
    Dim circleShape As shape
    Dim targetColor As Long
    Dim cellValue As Double
    Dim storeTextBox As shape

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Check if the text box "store" exists
    On Error Resume Next
    Set storeTextBox = pptSlide.Shapes("store")
    On Error GoTo 0

    If storeTextBox Is Nothing Then
        MsgBox "The 'store' text box was not found. Please run 'Create_balance_1' first.", vbExclamation
        Exit Sub
    End If

    ' Retrieve the stored value from the "store" text box
    cellValue = val(storeTextBox.TextFrame.textRange.text)
    Debug.Print "Retrieved value from 'store' text box: " & cellValue

    ' Ensure that the circle already exists on the slide
    On Error Resume Next
    Set circleShape = pptSlide.Shapes("BalanceCircle")
    On Error GoTo 0

    If circleShape Is Nothing Then
        MsgBox "The circle was not created. Please run 'Create_balance_2' first.", vbExclamation
        Exit Sub
    End If

    ' Adjust color based on percentage value
    If cellValue < 40 Then
        targetColor = RGB(228, 107, 127) ' Red
    ElseIf cellValue >= 40 And cellValue <= 70 Then
        targetColor = RGB(255, 172, 0) ' Orange
    Else
        targetColor = RGB(153, 208, 204) ' Light green
    End If

    ' Set the color of the circle
    circleShape.Fill.Solid
    circleShape.Fill.ForeColor.RGB = targetColor
    Debug.Print "Circle color set to: " & targetColor
End Sub

