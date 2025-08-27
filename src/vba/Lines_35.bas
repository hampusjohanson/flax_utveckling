Attribute VB_Name = "Lines_35"
Sub Lines_35()
' Select the "left_chart"
    Dim pptSlide As slide
    Dim leftChart As shape

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
    Set leftChart = pptSlide.Shapes("left_chart")
    On Error GoTo 0

    ' Check if "left_chart" exists
    If leftChart Is Nothing Then
        MsgBox "The chart named 'left_chart' was not found on this slide.", vbExclamation
        Exit Sub
    End If

    ' Select the "left_chart"
    leftChart.Select
End Sub

