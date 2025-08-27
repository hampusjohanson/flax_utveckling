Attribute VB_Name = "Examples_Custom_Ribbons"
' Global variabel för användarnamn
Public gUserInput As String


Public Sub ComboUserInput_Change(control As IRibbonControl, text As String)
    gUserInput = text ' Uppdatera den globala variabeln
    Debug.Print "Value updated in gUserInput: " & gUserInput ' Logga för att bekräfta uppdatering
    MsgBox "You selected: " & gUserInput, vbInformation ' Visa värdet
End Sub


Sub InitializeUserName()
    ' Kontrollera om användarnamnet redan är satt
    If Trim(gUserInput) = "" Then
        ' Fråga användaren efter användarnamn
        gUserInput = InputBox("Ange ditt användarnamn (endast första gången).", "Användarnamn")
        
        ' Om inget användarnamn anges, avbryt
        If Trim(gUserInput) = "" Then
            MsgBox "Inget användarnamn angavs. Makrot avbryts.", vbCritical
            Exit Sub
        End If
    End If
End Sub

Sub ProcessTopDriver()
    ' Ensure a value is selected before proceeding
    If selectedTopDriver = vbNullString Then
        MsgBox "No top driver selected!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Example usage: display the selected value
    MsgBox "Processing for: " & selectedTopDriver, vbInformation, "Top Driver"

    ' Add your logic here to process the selected value
End Sub

Sub CheckUserInputStatus()
    If gUserInput = "" Then
        MsgBox "gUserInput är inte aktiv. Användarnamnet har inte angetts.", vbExclamation
    Else
        MsgBox "gUserInput är aktiv. Användarnamnet är: " & gUserInput, vbInformation
    End If
End Sub

