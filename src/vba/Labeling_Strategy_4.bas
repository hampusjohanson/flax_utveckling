Attribute VB_Name = "Labeling_Strategy_4"
Sub Labeling_Strategy_4()
   On Error Resume Next ' Avoid breaking if a macro fails

    
    ' Steg 1: Rensa labels
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' Steg 2: Återställ labels
    Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

Application.Run "IdentifyAndMoveBottomFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "IdentifyAndMoveTopFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
Application.Run "IdentifyAndMoveLeftFlankLabels_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "IdentifyAndMoveRightFlankLabels_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

 Application.Run "IdentifyAndMoveLeftFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "IdentifyAndMoveRightFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
          Application.Run "IdentifyAndMoveBottomFlankLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
    
 
    ActiveWindow.ViewType = ppViewSlideSorter
    DoEvents
    ActiveWindow.ViewType = ppViewNormal
    DoEvents

    Application.Run "SaveLabelDistanceCounts"
  End Sub
