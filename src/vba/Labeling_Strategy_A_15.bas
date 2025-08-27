Attribute VB_Name = "Labeling_Strategy_A_15"
Sub Labeling_Strategy_15()
    
       Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
     Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
Application.Run "IdentifyAndMoveLeftFlankLabels_5"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels_5"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveLeftFlankLabels_20"
Application.Run "IdentifyAndMoveRightFlankLabels_20"
 
 
    
End Sub



