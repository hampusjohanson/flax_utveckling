Attribute VB_Name = "LD_1_2"
Sub LD_1_2()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you on slide LIE DETECTOR?", vbYesNo + vbQuestion, "Check Slide")
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "LD_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "LD_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
        
        Application.Run "ListMissingVisiblePoints_Final"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
           Application.Run "CloseChartExcel_MacSafe"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
        Application.Run "WaitOneSecond"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CloseChartExcel_MacSafe"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


    
End Sub

