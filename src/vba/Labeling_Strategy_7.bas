Attribute VB_Name = "Labeling_Strategy_7"
Sub Labeling_Strategy_7()
    On Error Resume Next ' Avoid breaking if a macro fails
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "IdentifyAndMoveTopFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      
      Application.Run "IdentifyAndMoveBottomFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
   Application.Run "IdentifyAndMoveLeftFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
 
      Application.Run "IdentifyAndMoveRightFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
        ' Aktivera ledande linjer
          lbl.ShowLeaderLines = True
End Sub



