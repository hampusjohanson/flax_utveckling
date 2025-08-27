Attribute VB_Name = "Export_macro_info"
Sub ExportAllMacros_PPT()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim exportFile As Integer
    Dim filePath As String

    filePath = Environ("USERPROFILE") & "\Desktop\All_PPT_Macros.txt"
    exportFile = FreeFile
    Open filePath For Output As exportFile

    Set vbProj = Application.VBE.ActiveVBProject

    For Each vbComp In vbProj.VBComponents
        Set codeMod = vbComp.CodeModule
        If codeMod.CountOfLines > 0 Then
            Print #exportFile, "----- Modul: " & vbComp.Name & " -----"
            Print #exportFile, codeMod.Lines(1, codeMod.CountOfLines)
            Print #exportFile, vbCrLf
        End If
    Next vbComp

    Close exportFile
    Debug.Print "Makron exporterade till: " & filePath
End Sub

