Attribute VB_Name = "Create_Balance_8"
Sub Create_balance_8()
    Dim pptSlide As slide
    Dim targetTable As table
    Dim textBoxShape As shape
    Dim text20 As String
    Dim text21 As String

    ' Reference the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Ensure the "TARGET" table exists
    On Error Resume Next
    Set targetTable = pptSlide.Shapes("TARGET").table
    On Error GoTo 0

    If targetTable Is Nothing Then
        MsgBox "The 'TARGET' table was not found. Please ensure the table exists.", vbExclamation
        Exit Sub
    End If

    ' === Extract text from TARGET table, columns 2 and 3, row 1, and convert to lowercase ===
    text20 = LCase(Trim(targetTable.cell(1, 2).shape.TextFrame.textRange.text))
    text21 = LCase(Trim(targetTable.cell(1, 3).shape.TextFrame.textRange.text))

    Debug.Print "Extracted Text 1: " & text20
    Debug.Print "Extracted Text 2: " & text21

    ' === Ensure the existing text box is found ===
    On Error Resume Next
    Set textBoxShape = pptSlide.Shapes("MyTextBox")
    On Error GoTo 0

    If textBoxShape Is Nothing Then
        MsgBox "'MyTextBox' text box not found. Please create the text box first.", vbExclamation
        Exit Sub
    End If

    ' === Update the text inside the existing text box ===
    textBoxShape.TextFrame.textRange.text = "= similarity between " & text20 & " and " & text21 & "."

    ' === Format the text in the text box ===
    With textBoxShape.TextFrame.textRange
        .Font.Name = "Arial"
        .Font.size = 10
        .Font.color = RGB(17, 21, 66) ' Dark blue color
    End With

    ' Enable WordWrap for the text box
    textBoxShape.TextFrame.WordWrap = msoTrue

    ' Ensure the text box is large enough to fit multiple lines if necessary
    textBoxShape.height = 2 * 28.35 ' Increase the height to allow WordWrap

    ' Debugging - Verify the content inside the text box
    Debug.Print "Text inside the 'MyTextBox': " & textBoxShape.TextFrame.textRange.text
End Sub

