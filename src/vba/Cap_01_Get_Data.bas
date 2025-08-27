Attribute VB_Name = "Cap_01_Get_Data"

Sub Cap_01_Get_Data()
    On Error Resume Next ' Avoid breaking if a macro fails
 Dim response As VbMsgBoxResult
    Dim pptSlide As slide
    Dim shp As shape
    Dim hasChart As Boolean

    ' Fråga användaren om de är på rätt slide
    response = MsgBox("Are you on CAPITALIZATION chart slide?", vbYesNo + vbQuestion, "Bekräfta")
    If response = vbNo Then Exit Sub

      Application.Run "Capital_111"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
       Application.Run "Mac_Cap_Axes"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    Application.Run "Mac_Cap_color"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
      Application.Run "Mac_Cap_Labels1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Mac_Cap_color"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     Application.Run "Mac_Cap_Remove_Crossings"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
     
     Application.Run "Mac_Cap_trendline"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


End Sub



