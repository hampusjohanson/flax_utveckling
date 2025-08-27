Attribute VB_Name = "Labeling_Strategy_A_13"
Sub Labeling_Strategy_13()
    On Error Resume Next ' Avoid breaking if a macro fails
    
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
Application.Run "IdentifyAndMoveLeftFlankLabels_30"
Application.Run "IdentifyAndMoveTopFlankLabels_"
Application.Run "IdentifyAndMoveRightFlankLabels_30"
 
 
    
End Sub


