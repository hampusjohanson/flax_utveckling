Attribute VB_Name = "Labeling_Strategy_3"
Sub Labeling_Strategy_3()
   On Error Resume Next ' Avoid breaking if a macro fails

    ' Steg 1: Rensa labels
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' Steg 2: Återställ labels
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    
    Application.Run "IdentifyAndMoveRightFlankLabels_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    Application.Run "IdentifyAndMoveLeftFlankLabels_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
  
  
    ActiveWindow.ViewType = ppViewSlideSorter
    DoEvents
    ActiveWindow.ViewType = ppViewNormal
    DoEvents
    Application.Run "SaveLabelDistanceCounts"
    DoEvents
    
End Sub

