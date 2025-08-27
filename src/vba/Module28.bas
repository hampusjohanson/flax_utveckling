Attribute VB_Name = "Module28"
Sub ExportAllMacros_PPT()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim i As Long
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
            For i = 1 To codeMod.CountOfLines
                Print #exportFile, codeMod.Lines(i, 1)
            Next i
            Print #exportFile, vbCrLf
        End If
    Next vbComp

    Close exportFile
    MsgBox "Makron exporterade till: " & filePath, vbInformation
End Sub

