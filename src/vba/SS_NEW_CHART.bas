Attribute VB_Name = "SS_NEW_CHART"
Sub SS_NEW_CHART()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub



