Attribute VB_Name = "LD_1_5"
Sub LD_1_5()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub


