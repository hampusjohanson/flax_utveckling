Attribute VB_Name = "Awareness_2a_Select_Brands"
Sub Awareness_Select_Brands()
    
    On Error Resume Next ' Avoid breaking if a macro fails

      Application.Run "Insert_Brand_Table"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    

End Sub





