Attribute VB_Name = "ToggleWrapLabel"
Public SelectedLanguage As String

Sub SetLanguageToSwedish()
    SelectedLanguage = "Swedish"
    MsgBox "Language set to Swedish.", vbInformation
End Sub

Sub SetLanguageToEnglish()
    SelectedLanguage = "English"
    MsgBox "Language set to English.", vbInformation
End Sub

