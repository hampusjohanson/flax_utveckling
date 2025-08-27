Attribute VB_Name = "Lines_Set_Markers"
Sub Lines_Set_Markers()
    On Error Resume Next ' Avoid breaking if a macro fails


    Application.Run "Lines_42"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

    Application.Run "SetMarkersForAllP"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents


End Sub



