Attribute VB_Name = "labeling_move_totals"
Sub labeling_move_total_1()
    On Error Resume Next ' Avoid breaking if a macro fails


    Application.Run "labeling_move_2"
    Application.Run "labeling_move_3_identify_leftright"
    Application.Run "labeling_move_4"
    Application.Run "labeling_move_5"
 
End Sub



Sub labeling_move_total_2()
    On Error Resume Next ' Avoid breaking if a macro fails

  
    Application.Run "labeling_move_2"
    Application.Run "labeling_move_3_identify_topbottom"
    Application.Run "labeling_move_4"
    Application.Run "labeling_move_5"
  
End Sub


