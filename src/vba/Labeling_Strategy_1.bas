Attribute VB_Name = "Labeling_Strategy_1"
Sub Labeling_Strategy_1()
    On Error Resume Next ' Avoid breaking if a macro fails

    ' Steg 1: Rensa labels
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' Steg 2: �terst�ll labels
    Application.Run "DataLabels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' Steg 3: V�xla vyer f�r att tvinga PowerPoint att rita om innan vi m�ter avst�nd
    ActiveWindow.ViewType = ppViewSlideSorter
    DoEvents
    ActiveWindow.ViewType = ppViewNormal
    DoEvents

    Application.Run "SaveLabelDistanceCounts"
  DoEvents
    
End Sub
