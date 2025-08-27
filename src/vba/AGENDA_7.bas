Attribute VB_Name = "AGENDA_7"
Sub AGENDA_7()
    Dim i As Integer, j As Integer
    Dim agendaText As String, chapterText As String
    Dim msg As String
    Dim missingCount As Integer, extraCount As Integer
    Dim chapterKey As Variant, agendaKey As Variant
    Dim agendaSplit As Variant
    Dim processedAgenda() As String, processedChapter() As String
    Dim agendaSize As Integer, chapterSize As Integer
    
    If agendaItems Is Nothing Or agendaItems.count = 0 Then
        MsgBox "agendaItems is empty!", vbExclamation, "Error"
        Exit Sub
    End If
    
    If chapterItems Is Nothing Or chapterItems.count = 0 Then
        MsgBox "chapterItems is empty!", vbExclamation, "Error"
        Exit Sub
    End If
    
    agendaSize = 0
    chapterSize = 0
    
    ' Process Agenda Items
    For i = 1 To agendaItems.count
        If InStr(agendaItems(i), Chr(10)) > 0 Then
            agendaSplit = Split(agendaItems(i), Chr(10)) ' Line feed split
        ElseIf InStr(agendaItems(i), Chr(13)) > 0 Then
            agendaSplit = Split(agendaItems(i), Chr(13)) ' Carriage return split
        Else
            agendaSplit = Array(agendaItems(i)) ' No split needed
        End If
        
        For j = LBound(agendaSplit) To UBound(agendaSplit)
            agendaText = Trim(agendaSplit(j))
            If agendaText <> "" Then
                ReDim Preserve processedAgenda(agendaSize)
                processedAgenda(agendaSize) = agendaText
                agendaSize = agendaSize + 1
            End If
        Next j
    Next i
    
    ' Process Chapter Items
    For i = 1 To chapterItems.count
        chapterText = Trim(chapterItems(i))
        If chapterText <> "" Then
            ReDim Preserve processedChapter(chapterSize)
            processedChapter(chapterSize) = chapterText
            chapterSize = chapterSize + 1
        End If
    Next i
    
    ' Compare Agenda and Chapter Items
    msg = "Kapitelbilder som inte finns agendan:" & vbCrLf
    missingCount = 0
    
    For i = 0 To chapterSize - 1
        If Not IsInArray(processedChapter(i), processedAgenda) Then
            msg = msg & "- " & processedChapter(i) & vbCrLf
            missingCount = missingCount + 1
        End If
    Next i
    
    If missingCount = 0 Then
        msg = msg & "None" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Agendapunkt som inte har kapitelbild:" & vbCrLf
    extraCount = 0
    
    For i = 0 To agendaSize - 1
        If Not IsInArray(processedAgenda(i), processedChapter) Then
            msg = msg & "- " & processedAgenda(i) & vbCrLf
            extraCount = extraCount + 1
        End If
    Next i
    
    If extraCount = 0 Then
        msg = msg & "None" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Agenda och kapitelbilder"
End Sub

Function IsInArray(value As String, arr() As String) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i)) = LCase(value) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

