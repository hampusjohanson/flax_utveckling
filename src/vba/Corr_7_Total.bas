Attribute VB_Name = "Corr_7_Total"


Sub Corr_7_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "Corr_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
   Application.Run "Corr_Bold_9999"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


Application.Run "Corr_Bold_99"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


End Sub





