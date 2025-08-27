Attribute VB_Name = "Lines_37"
Sub Lines_37()
    Dim pptSlide As slide
    Dim rightChart As shape

    ' Get the active slide
    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    On Error GoTo 0

    If pptSlide Is Nothing Then
        MsgBox "No active slide found. Please make sure you're in Normal View and try again.", vbExclamation
        Exit Sub
    End If

    ' Find the chart named "left_chart"
    On Error Resume Next
    Set rightChart = pptSlide.Shapes("right_chart")
    On Error GoTo 0

    ' Check if "left_chart" exists
    If rightChart Is Nothing Then
        MsgBox "The chart named 'right_chart' was not found on this slide.", vbExclamation
        Exit Sub
    End If

    rightChart.Select
End Sub


