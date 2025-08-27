Attribute VB_Name = "SP_00"
Sub SP_0()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "SP_01"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "SP_02"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


End Sub



