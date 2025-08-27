Attribute VB_Name = "Corr_2_Total"


Sub Corr_2_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "Corr_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "Corr_2_Text"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "Corr_Bold_9999"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


Application.Run "Corr_Bold_99"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


End Sub



