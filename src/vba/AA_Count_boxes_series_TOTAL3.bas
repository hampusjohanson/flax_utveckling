Attribute VB_Name = "AA_Count_boxes_series_TOTAL3"
Sub AA_Series_set_as_Total3()
    ' Check for "Rightie" and "Leftie"
    If Not AA_Check_Rightie_Leftie() Then Exit Sub

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "aware_smaller_height"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "AA_Series_set_as_TotalA"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
      
    Application.Run "AA_Series_set_as_Total2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "AA_Series_set_as_TotalA"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
End Sub

