Attribute VB_Name = "ApplyEyedropperColor_1"
Sub ApplyEyedropperColor(control As IRibbonControl)
    Dim shp As shape
    Dim colorRGB As Long

    ' Check if a shape is selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
    Else
        MsgBox "Please select a shape first.", vbExclamation, "No Shape Selected"
        Exit Sub
    End If

    ' Use AppleScript to activate macOS color picker
    Dim script As String
    script = "set pickedColor to choose color" & vbNewLine & _
             "return (item 1 of pickedColor as integer) * 65536 + (item 2 of pickedColor as integer) * 256 + (item 3 of pickedColor as integer)"

    ' Execute AppleScript and get color in RGB
    colorRGB = MacScript(script)

    ' Apply the picked color to the shape's fill
    shp.Fill.ForeColor.RGB = colorRGB

    ' Confirmation message
    MsgBox "Color applied!", vbInformation, "Success"
End Sub


