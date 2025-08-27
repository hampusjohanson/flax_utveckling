Attribute VB_Name = "Clear_Both"
Sub Clear_Both()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "Clear_Left"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "Clear_Right"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description


End Sub



