Attribute VB_Name = "Module22"
Sub FindMacro_Chart_Remove_series()
    Dim vbComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim lineText As String
    Dim targetMacroName As String
    Dim found As Boolean

    targetMacroName = "Chart_Remove_series"
    found = False

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Set codeMod = vbComp.CodeModule
        For i = 1 To codeMod.CountOfLines
            lineText = Trim(LCase(codeMod.Lines(i, 1)))
            If lineText Like "sub " & LCase(targetMacroName) & "*" Or _
               lineText Like "public sub " & LCase(targetMacroName) & "*" Then
                Debug.Print "Makrot '" & targetMacroName & "' finns i modul: " & vbComp.Name
                found = True
                Exit For
            End If
        Next i
        If found Then Exit For
    Next vbComp

    If Not found Then
        Debug.Print "Makrot '" & targetMacroName & "' hittades inte."
    End If
End Sub


