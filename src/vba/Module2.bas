Attribute VB_Name = "Module2"
Public flagSize As Double

Sub ComboSizeSelector_Change(control As IRibbonControl, text As String)
    ' Map size options to height values
    Select Case text
        Case "Extra Small"
            flagSize = 50
        Case "Small"
            flagSize = 100
        Case "Medium"
            flagSize = 150
        Case "Large"
            flagSize = 200
        Case "Extra Large"
            flagSize = 300
        Case Else
            flagSize = 150 ' Default to Medium if not set
    End Select

    ' Debugging: Print the selected size
    Debug.Print "Selected Size: " & text & ", Flag Height: " & flagSize
End Sub

Sub InsertSelectedFlag(control As IRibbonControl)
    Dim pptSlide As slide
    Dim hiddenSlide As slide
    Dim flagShape As shape
    Dim copiedFlag As shape

    ' Ensure a flag is selected
    If Len(Trim(selectedFlag)) = 0 Then
        MsgBox "No flag selected. Please choose a flag from the dropdown.", vbCritical
        Exit Sub
    End If

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Get the slide with flag images (last slide)
    Set hiddenSlide = ActivePresentation.Slides(ActivePresentation.Slides.count)

    ' Find the flag image by its name
    On Error Resume Next
    Set flagShape = hiddenSlide.Shapes(selectedFlag)
    On Error GoTo 0

    If flagShape Is Nothing Then
        MsgBox "Flag image for '" & selectedFlag & "' not found. Ensure the image exists and is named '" & selectedFlag & "'.", vbExclamation
        Exit Sub
    End If

    ' Copy the flag image and paste it on the current slide
    flagShape.Copy
    Set copiedFlag = pptSlide.Shapes.Paste(1)

    ' Adjust position and size based on user inputs
    With copiedFlag
        .LockAspectRatio = msoTrue
        .Top = 100 ' Default Top position
        .left = 100 ' Default Left position
        .height = IIf(flagSize > 0, flagSize, 150) ' Use selected size, default to Medium
    End With

    ' Confirm the insertion
    MsgBox "Flag '" & selectedFlag & "' has been inserted with size: " & flagSize, vbInformation
End Sub

