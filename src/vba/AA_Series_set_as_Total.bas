Attribute VB_Name = "AA_Series_set_as_Total"
Sub AA_Series_set_as_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "AA_Series_set_1_as_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "AA_Series_set_2_as_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

       Application.Run "AA_Series_set_3_as_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

End Sub


