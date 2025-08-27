Attribute VB_Name = "Labeling_Strategy_A_12"
Sub Labeling_Strategy_12()
    On Error Resume Next ' Avoid breaking if a macro fails
    
       Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
Application.Run "IdentifyAndMoveTopFlankLabels"
Application.Run "IdentifyAndMoveLeftFlankLabels"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels"
 
    
End Sub

