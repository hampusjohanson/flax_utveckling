Attribute VB_Name = "Table_Tool_Total"
Sub Table_Tool_Total()
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "Table_Tool_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
         Application.Run "Table_Tool_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "Table_Tool_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "Table_Tool_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       
    Application.Run "Table_Tool_5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
        Application.Run "Table_Tool_6"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
    
End Sub

