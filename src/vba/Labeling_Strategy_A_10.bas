Attribute VB_Name = "Labeling_Strategy_A_10"
Sub Labeling_Strategy_10()
  On Error Resume Next ' Avoid breaking if a macro fails
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
Application.Run "IdentifyAndMoveLeftFlankLabels"
Application.Run "IdentifyAndMoveTopFlankLabels"
Application.Run "IdentifyAndMoveBottomFlankLabels"
Application.Run "IdentifyAndMoveRightFlankLabels"
  

        ' Aktivera ledande linjer
          lbl.ShowLeaderLines = True
End Sub






