Attribute VB_Name = "Labeling_Strategy_2"
Sub Labeling_Strategy_2()

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
    
    ' Steg 4: Växla vyer för att tvinga PowerPoint att rita om innan vi mäter avstånd
    ActiveWindow.ViewType = ppViewSlideSorter
    DoEvents
    ActiveWindow.ViewType = ppViewNormal
    DoEvents

    ' Steg 5: Spara uppdaterade koordinater
    Application.Run "SaveLabelDistanceCounts"

End Sub
