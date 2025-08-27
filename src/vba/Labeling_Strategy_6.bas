Attribute VB_Name = "Labeling_Strategy_6"
Sub Labeling_Strategy_6()
    On Error Resume Next ' Avoid breaking if a macro fails
    
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AlignDataLabelsLeft"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "IdentifyAndMoveTopFlankLabels"
  
      
      Application.Run "IdentifyAndMoveBottomFlankLabels"
  
    
   Application.Run "IdentifyAndMoveLeftFlankLabels"
   
 
      Application.Run "AdjustFlankLabels2"
 
    
        ' Aktivera ledande linjer
          lbl.ShowLeaderLines = True
End Sub


