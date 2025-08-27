Attribute VB_Name = "Lines_38"
Sub Lines_38_Total()
'Update Weaker Side

    On Error Resume Next ' Avoid breaking if a macro fails

 Application.Run "Lines_34"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

      Application.Run "Lines_37"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "Lines_32"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    

End Sub




