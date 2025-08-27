Attribute VB_Name = "Labeling_Strategy_A_26"
Sub Labeling_Strategy_26()
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

   
      Application.Run "AdjustFlankLabels3"
    
End Sub



