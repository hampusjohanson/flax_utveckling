Attribute VB_Name = "Agenda_Total2"
Sub Agenda_Total2()

  Application.Run "Agenda_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

  Application.Run "Agenda_7"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
End Sub

