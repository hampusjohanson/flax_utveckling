Attribute VB_Name = "lines_total_run"
Sub Lines_Total_Run()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you at a slide with two LINJEDIAGRAM?", vbYesNo + vbQuestion, "Check Slide")
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "Lines_0"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_1a"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_1b"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_36_Total"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_38_Total"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_40"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_41"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_Brand_List"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Lines_42"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "SetMarkersForAllP"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
   Application.Run "Lines_10"
       Application.Run "Lines_11_12a"
    
      Application.Run "Adjust_Right_Chart_By_Axis_Settings"
    
End Sub

