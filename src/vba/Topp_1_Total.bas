Attribute VB_Name = "Topp_1_Total"
Sub Topp_1_Total()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you on den typen av drivkrafts-slide med FYRA KOLUMNER MED TIO RADER I VARJE?", vbYesNo + vbQuestion, "Check Slide")
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "Topp_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Topp_Update_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Topp_Update_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Topp_Update_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Topp_Update_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Topp_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
End Sub

