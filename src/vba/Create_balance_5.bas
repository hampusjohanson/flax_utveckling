Attribute VB_Name = "Create_balance_5"
Sub create_balance_5()
    Dim pptSlide As slide
    Dim storeTextBox As shape

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Check if the text box "store" exists
    On Error Resume Next
    Set storeTextBox = pptSlide.Shapes("store")
    On Error GoTo 0

    If storeTextBox Is Nothing Then
        MsgBox "The 'store' text box was not found.", vbExclamation
        Exit Sub
    End If

    ' Delete the "store" text box
    storeTextBox.Delete

    ' Debugging - Confirm deletion
    Debug.Print "'store' text box deleted."
End Sub

