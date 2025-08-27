Attribute VB_Name = "AA_Total_3series"
Sub AA_Total_3series()
    
    On Error Resume Next ' Avoid breaking if a macro fails


      Application.Run "AW_2_Remove_Series"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Insert_Brand_Table"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
          Application.Run "AA_Series_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
              Application.Run "AA_Series_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

              Application.Run "AA_Series_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "AA_Names"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
              Application.Run "DeleteBrandsTable"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


End Sub




