Attribute VB_Name = "Labeling_Count_Overlap_Dist"
Sub Labeling_Count_Overlap_Dist()
    ' Run the label distance check
    Application.Run "SaveLabelDistanceCounts"

    ' Display results in a message box
    MsgBox "This many distant errors: " & FarLabelCount & vbNewLine & vbNewLine & _
           "These labels are far off:" & vbNewLine & farLabels, vbInformation, "Label Distance Check"
End Sub


