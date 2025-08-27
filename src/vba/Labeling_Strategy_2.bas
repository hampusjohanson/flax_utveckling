Attribute VB_Name = "Labeling_Strategy_2"
Sub Labeling_Strategy_2()

    ' Steg 1: Rensa labels
    Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    ' Steg 2: �terst�ll labels
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    
Application.Run "IdentifyAndMoveRightFlankLabels_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    ' Steg 4: V�xla vyer f�r att tvinga PowerPoint att rita om innan vi m�ter avst�nd
    ActiveWindow.ViewType = ppViewSlideSorter
    DoEvents
    ActiveWindow.ViewType = ppViewNormal
    DoEvents

    ' Steg 5: Spara uppdaterade koordinater
    Application.Run "SaveLabelDistanceCounts"

End Sub
