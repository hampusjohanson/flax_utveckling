Attribute VB_Name = "SUB_TABLE_SCENARIO_2_Total"
Sub SUB_TABLE_SCENARIO_2_Total()
    On Error Resume Next ' Avoid breaking if a macro fails

      
        Application.Run "SUB_TABLE_SCENARIO_2_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
       
         Application.Run "SUB_TABLE_SCENARIO_2_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
       
      Application.Run "SUB_TABLE_SCENARIO_2_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
    
    
End Sub


