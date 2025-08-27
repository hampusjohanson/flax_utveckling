Attribute VB_Name = "AB_Total"
Sub AB_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "AB_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "AB_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

       Application.Run "AB_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
     Application.Run "AB_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

End Sub

