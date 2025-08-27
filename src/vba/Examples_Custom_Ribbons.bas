Attribute VB_Name = "Examples_Custom_Ribbons"
' Global variabel f�r anv�ndarnamn
Public gUserInput As String


Public Sub ComboUserInput_Change(control As IRibbonControl, text As String)
    gUserInput = text ' Uppdatera den globala variabeln
    Debug.Print "Value updated in gUserInput: " & gUserInput ' Logga f�r att bekr�fta uppdatering
    MsgBox "You selected: " & gUserInput, vbInformation ' Visa v�rdet
End Sub


Sub InitializeUserName()
    ' Kontrollera om anv�ndarnamnet redan �r satt
    If Trim(gUserInput) = "" Then
        ' Fr�ga anv�ndaren efter anv�ndarnamn
        gUserInput = InputBox("Ange ditt anv�ndarnamn (endast f�rsta g�ngen).", "Anv�ndarnamn")
        
        ' Om inget anv�ndarnamn anges, avbryt
        If Trim(gUserInput) = "" Then
            MsgBox "Inget anv�ndarnamn angavs. Makrot avbryts.", vbCritical
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
        MsgBox "gUserInput �r inte aktiv. Anv�ndarnamnet har inte angetts.", vbExclamation
    Else
        MsgBox "gUserInput �r aktiv. Anv�ndarnamnet �r: " & gUserInput, vbInformation
    End If
End Sub

