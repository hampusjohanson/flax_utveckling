Attribute VB_Name = "Labeling_Strategy_A_21"
Sub Labeling_Strategy_21()
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
  
      
   
   Application.Run "IdentifyAndMoveLeftFlankLabels_30"
   
      Application.Run "IdentifyAndMoveRightFlankLabels_30"
  Application.Run "IdentifyAndMoveTopFlankLabels"
   
      Application.Run "IdentifyAndMoveBottomFlankLabels"
    
End Sub



