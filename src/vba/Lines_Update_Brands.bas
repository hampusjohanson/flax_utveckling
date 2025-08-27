Attribute VB_Name = "Lines_Update_Brands"
Sub Lines_Update_Brands()
    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "Lines_11a"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

Application.Run "SetMarkersForAllP"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    
End Sub
