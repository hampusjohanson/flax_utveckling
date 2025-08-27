Attribute VB_Name = "Labeling_Strategy_A_24"
Sub Labeling_Strategy_24()
    
       Application.Run "DeleteAllDataLabels"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
 Application.Run "AdjustFlankLabels2"
  
     Application.Run "AlignDataLabelsLeft"
 
     Application.Run "labeling_move_total_1"
    Application.Run "labeling_move_total_2"
End Sub





