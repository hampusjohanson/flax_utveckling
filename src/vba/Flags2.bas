Attribute VB_Name = "Flags2"
Sub SetFlag_Global(control As IRibbonControl)
    selectedFlag = "Global"
    InsertFlag_Main control
End Sub

Sub SetFlag_Japan(control As IRibbonControl)
    selectedFlag = "Japan"
    InsertFlag_Main control
End Sub

Sub SetFlag_India(control As IRibbonControl)
    selectedFlag = "India"
    InsertFlag_Main control
End Sub

Sub SetFlag_Australia(control As IRibbonControl)
    selectedFlag = "Australia"
    InsertFlag_Main control
End Sub

Sub SetFlag_Italy(control As IRibbonControl)
    selectedFlag = "Italy"
    InsertFlag_Main control
End Sub

Sub SetFlag_France(control As IRibbonControl)
    selectedFlag = "France"
    InsertFlag_Main control
End Sub

Sub SetFlag_China(control As IRibbonControl)
    selectedFlag = "China"
    InsertFlag_Main control
End Sub

Sub SetFlag_Brazil(control As IRibbonControl)
    selectedFlag = "Brazil"
    InsertFlag_Main control
End Sub

Sub SetFlag_Argentina(control As IRibbonControl)
    selectedFlag = "Argentina"
    InsertFlag_Main control
End Sub

Sub SetFlag_Netherlands(control As IRibbonControl)
    selectedFlag = "Netherlands"
    InsertFlag_Main control
End Sub

Sub SetFlag_Turkey(control As IRibbonControl)
    selectedFlag = "Turkey"
    InsertFlag_Main control
End Sub

Sub SetFlag_Portugal(control As IRibbonControl)
    selectedFlag = "Portugal"
    InsertFlag_Main control
End Sub

Sub SetFlag_Spain(control As IRibbonControl)
    selectedFlag = "Spain"
    InsertFlag_Main control
End Sub

Sub SetFlag_Belgium(control As IRibbonControl)
    selectedFlag = "Belgium"
    InsertFlag_Main control
End Sub

Sub SetFlag_Germany(control As IRibbonControl)
    selectedFlag = "Germany"
    InsertFlag_Main control
End Sub

Sub SetFlag_UnitedStates(control As IRibbonControl)
    selectedFlag = "United States"
    InsertFlag_Main control
End Sub

Sub SetFlag_UK(control As IRibbonControl)
    selectedFlag = "UK"
    InsertFlag_Main control
End Sub

Sub SetFlag_Denmark(control As IRibbonControl)
    selectedFlag = "Denmark"
    InsertFlag_Main control
End Sub

Sub SetFlag_Finland(control As IRibbonControl)
    selectedFlag = "Finland"
    InsertFlag_Main control
End Sub

Sub SetFlag_Norway(control As IRibbonControl)
    selectedFlag = "Norway"
    InsertFlag_Main control
End Sub

Sub SetFlag_Sweden(control As IRibbonControl)
    selectedFlag = "Sweden"
    InsertFlag_Main control
End Sub


Sub InsertFlag_Main(control As IRibbonControl)
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
        .Top = centerTop - (.height / 2) ' Adjust Top to keep center consistent
        .left = centerLeft - (.width / 2) ' Adjust Left to keep center consistent
    End With

End Sub

