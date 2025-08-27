Attribute VB_Name = "SP_Total"
Sub SP_Total()
    On Error Resume Next ' Avoid breaking if a macro fails

    Dim response As VbMsgBoxResult
    Dim pptSlide As slide
    Dim shp As shape
    Dim hasChart As Boolean

    ' Fråga användaren om de är på rätt slide
    response = MsgBox("Are you on SALES PREMIUM chart slide?", vbYesNo + vbQuestion, "Bekräfta")
    If response = vbNo Then Exit Sub

    ' Run SP_1
    Application.Run "SP_1"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

    ' Run SP_2
    Application.Run "SP_2"
    If Err.Number <> 0 Then MsgBox "Error in SP_2: " & Err.Description
    DoEvents

    '
    Application.Run "SP_3"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
    
   
    Application.Run "Mac_Cap_Labels1"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents


End Sub

