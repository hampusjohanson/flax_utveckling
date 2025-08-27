Attribute VB_Name = "Lines_36"
Sub Lines_36_Total()
'Update Stronger Side
    
    On Error Resume Next ' Avoid breaking if a macro fails

 Application.Run "Lines_34"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

      Application.Run "Lines_35"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "Lines_31"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    

End Sub



