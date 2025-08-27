Attribute VB_Name = "Lines_2_Input_Axis"
Sub Lines_2()
   'SetVerticalAxisWithMapping
   Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim userMax As String
    Dim userMin As String
    Dim maxValue As Double
    Dim minValue As Double

    ' Check if a shape is selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Markera diagram tack.", vbExclamation
        Exit Sub
    End If

    ' Check if the selected shape is a chart
    Set chartShape = ActiveWindow.Selection.ShapeRange(1)
    If Not chartShape.hasChart Then
        MsgBox "Markera ett giltigt diagram tack.", vbExclamation
        Exit Sub
    End If

    ' Get the chart object
    Set chartObject = chartShape.chart

    ' Prompt the user for the maximum value
    userMax = InputBox("Vid vilken drivkraft (siffra) ska diagrammet börja?", "Set Axis Max")
    ' Validate input
    If Not IsNumeric(userMax) Or val(userMax) < 1 Or val(userMax) > 50 Then
        MsgBox "Ange ett giltigt värde mellan 1 och 50.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt the user for the minimum value
    userMin = InputBox("Vilken drivkraft (siffra) ska vara längst ned?", "Set Axis Min")
    ' Validate input
    If Not IsNumeric(userMin) Or val(userMin) < 1 Or val(userMin) > 50 Then
        MsgBox "Ange ett giltigt värde mellan 1 och 50.", vbExclamation
        Exit Sub
    End If

    ' Convert inputs using the mapping
    maxValue = 51 - val(userMax) ' Map 50 = 1, 49 = 2, etc.
    minValue = 51 - val(userMin)

    ' Ensure max > min
    If maxValue <= minValue Then
        MsgBox "Max-värdet måste vara större än min-värdet.", vbExclamation
        Exit Sub
    End If

    ' Set the Min and Max values for the vertical axis
    With chartObject.Axes(xlValue)
        .MinimumScale = minValue
        .MaximumScale = maxValue
    End With

  
End Sub

