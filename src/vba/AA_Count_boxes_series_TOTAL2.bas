Attribute VB_Name = "AA_Count_boxes_series_TOTAL2"




Sub AA_Series_set_as_Total2()
    On Error Resume Next ' Avoid breaking if a macro fails

Application.Run "AA_rightest_leftie"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      
      Application.Run "AA_Count_boxes_series"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "AA_Count_boxes_series_A9"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
  

    Application.Run "AA_Count_boxes_series_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    

    Application.Run "AA_Count_boxes_series_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

End Sub


