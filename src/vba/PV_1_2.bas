Attribute VB_Name = "PV_1_2"
Sub PV_1_2()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "PV_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

         Application.Run "PV_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Mac_Cap_Remove_Crossings"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Mac_Cap_trendline"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    


End Sub





