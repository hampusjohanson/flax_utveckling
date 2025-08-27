Attribute VB_Name = "Labeling_Strategy_A_27"
Sub Labeling_Strategy_27()
    
 Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "IdentifyAndMoveLeftFlankLabels"
Application.Run "IdentifyAndMoveTopFlankLabels"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels"
   
   
      Application.Run "IdentifyAndMoveBottomFlankLabels"
    
    
 Application.Run "AdjustFlankLabels2"
End Sub


