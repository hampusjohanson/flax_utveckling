Attribute VB_Name = "Cap_00_Get_Table"


Sub Cap_00_Get_Table()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "SP_01"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "Capital_0"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

End Sub


