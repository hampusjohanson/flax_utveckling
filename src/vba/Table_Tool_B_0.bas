Attribute VB_Name = "Table_Tool_B_0"
Sub Table_Tool_B_0()
    ' Check if a text box is selected
    If ActiveWindow.Selection.Type <> ppSelectionText Then
        MsgBox "Please select a text box. The current macro chain will be terminated.", vbExclamation
        Exit Sub ' Stop further execution if no text box is selected
    End If
    
    ' Proceed with the macro chain if a text box is selected
    Debug.Print "Text box selected. Proceeding with macro chain..."

    ' If a text box is selected, we can call the next macro in the chain, for example:
    ' Call Table_Tool_B_1
End Sub

