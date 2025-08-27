Attribute VB_Name = "Create_Balance_Total"
Sub Create_Balance_Total2()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "Create_Balance_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

      Application.Run "Create_Balance_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Create_Balance_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

  Application.Run "Create_Balance_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

  Application.Run "Create_Balance_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
  Application.Run "Create_Balance_6"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
          Application.Run "Create_Balance_10"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
  Application.Run "Create_Balance_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Create_Balance_8"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    

End Sub



