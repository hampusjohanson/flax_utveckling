Attribute VB_Name = "Labeling_Strategy_A_18"
Sub Labeling_Strategy_18()
    
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
  
      Application.Run "IdentifyAndMoveTopFlankLabels"
   
      Application.Run "IdentifyAndMoveBottomFlankLabels"
   
   Application.Run "IdentifyAndMoveLeftFlankLabels_5"
   
      Application.Run "IdentifyAndMoveRightFlankLabels_5"
  
    
End Sub



