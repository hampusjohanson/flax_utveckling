Attribute VB_Name = "Background_1_Total"




Sub Background_1_Total()
    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "Background_1_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      
    Application.Run "Background_1_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "Background_1_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Background_1_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Background_1_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Background_1_6"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "Background_1_9"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Background_1_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
  
    Application.Run "Background_1_8"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
  
  
End Sub


