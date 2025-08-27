Attribute VB_Name = "Labeling_Strategy_A_11"



Sub Labeling_Strategy_11()
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
Application.Run "IdentifyAndMoveBottomFlankLabels_30"
Application.Run "IdentifyAndMoveRightFlankLabels_30"
  

        ' Aktivera ledande linjer
          lbl.ShowLeaderLines = True
End Sub







