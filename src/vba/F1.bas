Attribute VB_Name = "F1"
Sub Black_text_to_blue()
    Dim slide As slide
    Dim shape As shape
    Dim textRange As textRange
    Dim userResponse As VbMsgBoxResult
    Dim TotalCount As Long
    Dim i As Integer
    Dim j As Integer
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim textHasSearchColor As Boolean
    
    TotalCount = 0
    
    ' Ange färger för sök, ersätt och vit (att ignorera)
    Dim searchColorLowerBound As Long
    Dim searchColorUpperBound As Long
    Dim replaceColor As Long
    Dim ignoreColor As Long
    
    searchColorLowerBound = RGB(0, 0, 0) ' Minsta "svarta" nyans
    searchColorUpperBound = RGB(50, 50, 50) ' Största "svarta" nyans
    replaceColor = RGB(17, 21, 66) ' Ersätt med hex #111542 (blå färg)
    ignoreColor = RGB(255, 255, 255) ' Ignorera vit text (RGB #FFFFFF)
    
    On Error Resume Next ' Hoppa över problematiska objekt
    
    ' Loopa igenom alla slides i presentationen
    For Each slide In ActivePresentation.Slides
        ' Loopa igenom alla former på varje slide
        For Each shape In slide.Shapes
            
            ' 1. Hantera textrutor och former med text
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    textHasSearchColor = False
                    Set textRange = shape.TextFrame.textRange
                    
                    ' Kontrollera om texten är inom det svarta intervallet (och ignorera vit text)
                    For i = 1 To textRange.Characters.count
                        If IsColorInRange(textRange.Characters(i).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                           textRange.Characters(i).Font.color.RGB <> ignoreColor Then
                            textHasSearchColor = True
                            Exit For
                        End If
                    Next i
                    
                    ' Fråga för hela textrutan om färgen hittas
                    If textHasSearchColor Then
                        slide.Select
                        shape.Select
                        
                        userResponse = MsgBox("Svart text hittad på Slide " & slide.SlideIndex & "." & vbCrLf & _
                        "Vill du byta textfärg till blått?", vbYesNoCancel + vbQuestion, "Bekräfta färgändring")
                        
                        If userResponse = vbYes Then
                            For i = 1 To textRange.Characters.count
                                If IsColorInRange(textRange.Characters(i).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                                   textRange.Characters(i).Font.color.RGB <> ignoreColor Then
                                    textRange.Characters(i).Font.color.RGB = replaceColor
                                End If
                            Next i
                            TotalCount = TotalCount + 1
                        ElseIf userResponse = vbCancel Then
                            MsgBox "Makrot avbröts av användaren.", vbExclamation, "Avbrutet"
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            ' 2. Hantera tabeller som helhet
            If shape.HasTable Then
                textHasSearchColor = False
                rowCount = shape.table.Rows.count
                colCount = shape.table.Columns.count
                
                ' Kontrollera om någon cell har sökfärgen (och ignorera vit text)
                For i = 1 To rowCount
                    For j = 1 To colCount
                        If shape.table.cell(i, j).shape.HasTextFrame Then
                            If shape.table.cell(i, j).shape.TextFrame.HasText Then
                                Set textRange = shape.table.cell(i, j).shape.TextFrame.textRange
                                
                                For k = 1 To textRange.Characters.count
                                    If IsColorInRange(textRange.Characters(k).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                                       textRange.Characters(k).Font.color.RGB <> ignoreColor Then
                                        textHasSearchColor = True
                                        Exit For
                                    End If
                                Next k
                            End If
                        End If
                    Next j
                    If textHasSearchColor Then Exit For
                Next i
                
                ' Fråga för hela tabellen om färgen hittas
                If textHasSearchColor Then
                    slide.Select
                    shape.Select
                    
                    userResponse = MsgBox("Svart text hittad i en tabell på Slide " & slide.SlideIndex & "." & vbCrLf & _
                    "Vill du byta svart text till blått?", vbYesNoCancel + vbQuestion, "Bekräfta färgändring")
                    
                    If userResponse = vbYes Then
                        For i = 1 To rowCount
                            For j = 1 To colCount
                                If shape.table.cell(i, j).shape.HasTextFrame Then
                                    If shape.table.cell(i, j).shape.TextFrame.HasText Then
                                        Set textRange = shape.table.cell(i, j).shape.TextFrame.textRange
                                        For k = 1 To textRange.Characters.count
                                            If IsColorInRange(textRange.Characters(k).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                                               textRange.Characters(k).Font.color.RGB <> ignoreColor Then
                                                textRange.Characters(k).Font.color.RGB = replaceColor
                                            End If
                                        Next k
                                    End If
                                End If
                            Next j
                        Next i
                        TotalCount = TotalCount + 1
                    ElseIf userResponse = vbCancel Then
                        MsgBox "Makrot avbröts av användaren.", vbExclamation, "Avbrutet"
                        Exit Sub
                    End If
                End If
            End If
        Next shape
    Next slide
    
    On Error GoTo 0 ' Återställ normal felhantering
    
    ' Bekräftelsemeddelande
    MsgBox "Du bytte färg på " & TotalCount & " textrutor, tabeller och former från svart till blått.", vbInformation, "Färgändringar slutförda"
End Sub

' Funktion för att kontrollera om en färg är inom det angivna intervallet
Function IsColorInRange(color As Long, lowerBound As Long, upperBound As Long) As Boolean
    If color >= lowerBound And color <= upperBound Then
        IsColorInRange = True
    Else
        IsColorInRange = False
    End If
End Function







