Attribute VB_Name = "labeling_most_X4a"
Sub labeling_most_X4()
    Dim moveX As Single, moveY As Single

    ' Set movement values (adjust as needed)
    moveX = -1   ' Move left
    moveY = 0   ' Move up

    ' Ensure a label has been stored
    If LabelToMove Is Nothing Then
        MsgBox "No label has been stored for movement. Run the previous macro first.", vbExclamation
        Exit Sub
    End If

    ' Move the stored label
    LabelToMove.left = LabelToMove.left + moveX
    LabelToMove.Top = LabelToMove.Top + moveY

    ' Confirm movement
    Debug.Print "Moved stored label [" & LabelToMove.text & "] by X: " & moveX & ", Y: " & moveY
  
End Sub


