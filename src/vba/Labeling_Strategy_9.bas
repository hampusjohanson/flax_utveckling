Attribute VB_Name = "Labeling_Strategy_9"
Sub Labeling_Strategy_9()
  On Error Resume Next ' Avoid breaking if a macro fails
    
       Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

 

Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels"
    
End Sub




