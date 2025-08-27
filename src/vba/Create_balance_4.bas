Attribute VB_Name = "Create_balance_4"
Sub Create_balance_4()
    Dim pptSlide As slide
    Dim storeTextBox As shape
    Dim balanceCircle As shape
    Dim cellValue As Double

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

    ' Ensure that the "BalanceCircle" exists
    On Error Resume Next
    Set balanceCircle = pptSlide.Shapes("BalanceCircle")
    On Error GoTo 0

    If balanceCircle Is Nothing Then
        MsgBox "The 'BalanceCircle' was not found. Please run 'Create_balance_2' first.", vbExclamation
        Exit Sub
    End If

    ' Set the value inside the circle with "%" appended
    balanceCircle.TextFrame.textRange.text = cellValue & "%"

    ' Format the text inside the circle
    With balanceCircle.TextFrame.textRange
        .Font.Name = "Arial"  ' Set the font to Arial
        .Font.size = 16       ' Set the font size to 16
        .Font.Bold = msoTrue  ' Set the font to bold
        .ParagraphFormat.Alignment = ppAlignCenter ' Center the text inside the circle
    End With

    ' Adjust the left and right margins for the text inside the circle
    With balanceCircle.TextFrame
        .MarginLeft = 0    ' Adjust the left margin (you can adjust the value as needed)
        .MarginRight = 0   ' Adjust the right margin (you can adjust the value as needed)
    End With

    ' Debugging - Verify the value set inside the circle
    Debug.Print "Text set inside 'BalanceCircle': " & balanceCircle.TextFrame.textRange.text
End Sub

