Attribute VB_Name = "Module8"
Sub Toppa_Clean()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "ToppA_0"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "ToppA_01"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

          Application.Run "ToppA_02"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents



End Sub



