Attribute VB_Name = "AW_A6_NEW_CHART"
Sub AW_A6_NEW_CHART()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub




