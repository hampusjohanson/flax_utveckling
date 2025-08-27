Attribute VB_Name = "Lines_Legend_Total_10"


  Sub Lines_Legend_New_Total_10()

    On Error Resume Next ' skyddar mot enstaka missar utan att bryta hela
    Application.Run "Lines_Legend_Delete_Tables"
  Application.Run "Lines_Legend_New_1"
   Application.Run "Lines_Legend_New_2"
     Application.Run "Lines_Legend_New_3"
        Application.Run "Lines_Legend_New_4"
           Application.Run "Lines_Legend_New_5"
              Application.Run "Lines_Legend_New_6"
                 Application.Run "Lines_Legend_New_7"
                    Application.Run "Lines_Legend_New_8"
                    
        
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide

    On Error Resume Next
    pptSlide.Shapes("Brand_List_2").Delete
    On Error GoTo 0

    Debug.Print "Eventuella Brand_List-tabeller borttagna från sliden."
End Sub




