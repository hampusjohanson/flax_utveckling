Attribute VB_Name = "Scatter_fix_Update_data"
 Sub Scatter_fix_Update_data()
    On Error Resume Next ' Avoid breaking if a macro fails
    
    Application.Run "Scatter_fix_3"
 Application.Run "Scatter_fix_1"
  Application.Run "Scatter_fix_4"
   Application.Run "Scatter_fix_5"
   Application.Run "Scatter_fix_2"
      Application.Run "Scatter_fix_6"
End Sub

