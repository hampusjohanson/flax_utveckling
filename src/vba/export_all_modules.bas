Attribute VB_Name = "export_all_modules"
Sub ExportAllModules()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim exportPath As String

    ' Sätt exportmapp till projektets src\vba
    exportPath = "C:\Local\Flax utveckling\src\vba\"

    ' Referens till aktiva VBA-projektet
    Set vbProj = Application.VBE.ActiveVBProject

    ' Loopa igenom alla komponenter (moduler, klasser, userforms)
    For Each vbComp In vbProj.VBComponents
        ' Exportera standardmoduler och klassmoduler
        If vbComp.Type = 1 Or vbComp.Type = 2 Then
            vbComp.Export exportPath & vbComp.Name & ".bas"
        End If
    Next vbComp

    ' Rensa upp
    Set vbComp = Nothing
    Set vbProj = Nothing

    MsgBox "Alla moduler har exporterats till src\\vba.", vbInformation
End Sub
