Attribute VB_Name = "AW_9C"
Sub AW_9C()

  Application.Run "AddTransparentAndGreenOverlayFromExcel"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

  Application.Run "SetLabelToSumOfTwoBottomSeries"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
    
  Application.Run "ResizeAndRepositionChart"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub


