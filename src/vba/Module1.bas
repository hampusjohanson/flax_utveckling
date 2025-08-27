Attribute VB_Name = "Module1"
Public selectedFlag As String
Public flagSize As Double

Sub ComboFlagSelector_Change(control As IRibbonControl, text As String)
    ' Store the selected country
    selectedFlag = text
    Debug.Print "Selected Country: " & selectedFlag
End Sub

Sub ComboSizeSelector_Change(control As IRibbonControl, text As String)
    ' Map size options to height values
    Select Case text
        Case "Extra Small"
            flagSize = 22.6772 ' 0.8 cm
        Case "Small"
            flagSize = 34.0158 ' 1.2 cm
        Case "Medium"
            flagSize = 45.3544 ' 1.6 cm
        Case "Large"
            flagSize = 68.0316 ' 2.4 cm
        Case "Extra Large"
            flagSize = 90.7088 ' 3.2 cm
        Case Else
            flagSize = 45.3544 ' Default to Medium
    End Select
    Debug.Print "Selected Size: " & text & ", Flag Height: " & flagSize
End Sub
Sub InsertSelectedFlag(control As IRibbonControl)
    Dim pptSlide As slide
    Dim flagSlide As slide
    Dim flagShape As shape
    Dim copiedFlag As shape
    Dim slideFound As Boolean
    Dim centerTop As Double
    Dim centerLeft As Double

    ' Ensure a country is selected
    If Len(Trim(selectedFlag)) = 0 Then
        MsgBox "No country selected. Please choose a country from the dropdown.", vbCritical
        Exit Sub
    End If

    ' Define the center position (adjust these as needed for your layout)
    centerTop = 1.7 * 28.3465 ' Center vertical position in points (1.7 cm)
    centerLeft = 31.41 * 28.3465 ' Center horizontal position in points (31.41 cm)

    ' Find the slide with the headline "Flags"
    slideFound = False
    For Each pptSlide In ActivePresentation.Slides
        If pptSlide.Shapes.HasTitle Then
            If pptSlide.Shapes.title.TextFrame.textRange.text = "Flags" Then
                Set flagSlide = pptSlide
                slideFound = True
                Exit For
            End If
        End If
    Next pptSlide

    If Not slideFound Then
        MsgBox "Slide with the title 'Flags' not found. Please add a slide titled 'Flags' with the flag images.", vbCritical
        Exit Sub
    End If

    ' Find the flag image by its name on the "Flags" slide
    On Error Resume Next
    Set flagShape = flagSlide.Shapes(selectedFlag) ' selectedFlag must match the image name
    On Error GoTo 0

    If flagShape Is Nothing Then
        MsgBox "Flag image for '" & selectedFlag & "' not found on the 'Flags' slide. Ensure the image exists and is named '" & selectedFlag & "'.", vbExclamation
        Exit Sub
    End If

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Copy the flag image and paste it on the current slide
    flagShape.Copy
    Set copiedFlag = pptSlide.Shapes.Paste(1)

    ' Adjust position and size to keep the center consistent
    With copiedFlag
        .LockAspectRatio = msoTrue
        .height = IIf(flagSize > 0, flagSize, 45.3544) ' Use selected size, default to Medium
        .Top = centerTop - (.height / 2) ' Adjust Top to keep center consistent
        .left = centerLeft - (.width / 2) ' Adjust Left to keep center consistent
    End With

End Sub

