Attribute VB_Name = "Corr_3_Total"


Sub Corr_3_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "Corr_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "Corr_3_Text"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
   Application.Run "Corr_Bold_9999"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


Application.Run "Corr_Bold_99"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

End Sub



