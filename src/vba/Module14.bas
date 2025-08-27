Attribute VB_Name = "Module14"
Sub ImportAllModules()
    Dim vbProj As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importPath As String

    ' Set the folder path to the user's desktop
    importPath = Environ("USERPROFILE") & "\Desktop\"

    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(importPath)

    ' Reference the active VBA project
    Set vbProj = Application.VBE.ActiveVBProject

    ' Loop through all .bas files on the desktop
    For Each file In folder.Files
        If LCase(right(file.Name, 4)) = ".bas" Then
            vbProj.VBComponents.Import file.Path
        End If
    Next file

    ' Cleanup
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    Set vbProj = Nothing

    MsgBox "All .bas modules from desktop have been imported.", vbInformation
End Sub

