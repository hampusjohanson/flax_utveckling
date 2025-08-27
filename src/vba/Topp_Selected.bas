Attribute VB_Name = "Topp_Selected"
Public selectedTopDriver As String



' Callback for comboUserInput_topdriver
Sub ComboUserInput_Change_topdriver(control As IRibbonControl, text As String)
    ' Store the selected value in the global variable
    selectedTopDriver = text
    ' Debug output for validation
    Debug.Print "Selected Top Driver: " & selectedTopDriver
End Sub




