Attribute VB_Name = "Labeling_Strategy_A_20"
Sub Labeling_Strategy_20()
    On Error Resume Next ' Avoid breaking if a macro fails
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
  
      Application.Run "IdentifyAndMoveTopFlankLabels"
   
      Application.Run "IdentifyAndMoveBottomFlankLabels"
   
   Application.Run "IdentifyAndMoveLeftFlankLabels_20"
   
      Application.Run "IdentifyAndMoveRightFlankLabels_20"
  
    
End Sub

