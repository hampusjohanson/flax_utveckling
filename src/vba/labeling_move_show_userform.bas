Attribute VB_Name = "labeling_move_show_userform"
Sub ShowMoveForm()
    ' Ensure a label has been stored
    If LabelToMove Is Nothing Then
        MsgBox "No label has been stored for movement. Run the previous macro first.", vbExclamation
        Exit Sub
    End If

    ' Show the movement form
    MoveLabelForm.Show
End Sub

