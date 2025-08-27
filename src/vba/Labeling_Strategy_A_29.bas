Attribute VB_Name = "Labeling_Strategy_A_29"
Sub Labeling_Strategy_29()
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
  
      Application.Run "IdentifyAndMoveLeftFlankLabels"
Application.Run "IdentifyAndMoveTopFlankLabels"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels"
   
   Application.Run "IdentifyAndMoveLeftFlankLabels_30"
   
      Application.Run "IdentifyAndMoveRightFlankLabels_30"
  Application.Run "IdentifyAndMoveTopFlankLabels"
   
      Application.Run "IdentifyAndMoveBottomFlankLabels"
    
End Sub







