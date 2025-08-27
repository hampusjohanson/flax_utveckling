Attribute VB_Name = "Module15"
Sub AA_Series_11()
    Dim pptSlide As slide
    Dim newTextBox As shape

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Create a new text box with the specified dimensions and position
    Set newTextBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 3.54 * 28.35, 5.14 * 28.35, 1.73 * 28.35, 0.94 * 28.35)

    ' Set name for reference (optional)
    newTextBox.Name = "New_TextBox"

    ' Set default text
    newTextBox.TextFrame.textRange.text = "Your Text Here"

    ' Ensure text is visible and fits within the box
    newTextBox.TextFrame.AutoSize = ppAutoSizeNone

    MsgBox "New text box created with specified size and position.", vbInformation
End Sub

