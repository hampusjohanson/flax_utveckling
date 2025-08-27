Attribute VB_Name = "SUB_TABLE_SCENARIO_1_Total"
Sub SUB_TABLE_SCENARIO_1_Total()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "SUB_TABLE_SCENARIO_1a"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "SUB_TABLE_SCENARIO_1b"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
   Application.Run "SUB_TABLE_SCENARIO_1c"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "SUB_TABLE_SCENARIO_1d"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
End Sub


