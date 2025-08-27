Attribute VB_Name = "Lines_Set_Line_Size"
Sub Lines_Set_Line_Size()
    On Error Resume Next ' Avoid breaking if a macro fails




    Application.Run "Lines_3"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents
    
    
    Application.Run "Lines_7"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

 Application.Run "Lines_Set_Markers"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents


End Sub




