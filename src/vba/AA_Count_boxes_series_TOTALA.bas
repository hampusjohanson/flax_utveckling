Attribute VB_Name = "AA_Count_boxes_series_TOTALA"




Sub AA_Series_set_as_TotalA()
    On Error Resume Next ' Avoid breaking if a macro fails

Application.Run "AA_rightest_leftie"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      
      Application.Run "AA_Count_boxes_series"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "AA_Count_boxes_series_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

       Application.Run "AA_Count_boxes_series_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
           Application.Run "AA_Count_boxes_series_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

     Application.Run "AA_Count_boxes_series_A1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A6"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
     Application.Run "AA_Count_boxes_series_A8"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

End Sub

