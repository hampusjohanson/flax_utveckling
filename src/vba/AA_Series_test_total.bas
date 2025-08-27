Attribute VB_Name = "AA_Series_test_total"
Sub AA_Series_test_total()
    On Error Resume Next ' Avoid breaking if a macro fails


    
          Application.Run "AA_Series_6"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


          Application.Run "AA_Series_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents



Application.Run "AA_Series_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
Application.Run "AA_Series_8"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
 
 
Application.Run "AA_Series_9"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
Application.Run "AA_Series_10"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
 
End Sub



