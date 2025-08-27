Attribute VB_Name = "Table_Tool_B_Total"
Sub Table_Tool_B_Total()
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "Table_Tool_B_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
         Application.Run "Table_Tool_B_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
         Application.Run "Table_Tool_B_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
    
End Sub



