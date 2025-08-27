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
    
    ' Ange f�rger f�r s�k, ers�tt och vit (att ignorera)
    Dim searchColorLowerBound As Long
    Dim searchColorUpperBound As Long
    Dim replaceColor As Long
    Dim ignoreColor As Long
    
    searchColorLowerBound = RGB(0, 0, 0) ' Minsta "svarta" nyans
    searchColorUpperBound = RGB(50, 50, 50) ' St�rsta "svarta" nyans
    replaceColor = RGB(17, 21, 66) ' Ers�tt med hex #111542 (bl� f�rg)
    ignoreColor = RGB(255, 255, 255) ' Ignorera vit text (RGB #FFFFFF)
    
    On Error Resume Next ' Hoppa �ver problematiska objekt
    
    ' Loopa igenom alla slides i presentationen
    For Each slide In ActivePresentation.Slides
        ' Loopa igenom alla former p� varje slide
        For Each shape In slide.Shapes
            
            ' 1. Hantera textrutor och former med text
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    textHasSearchColor = False
                    Set textRange = shape.TextFrame.textRange
                    
                    ' Kontrollera om texten �r inom det svarta intervallet (och ignorera vit text)
                    For i = 1 To textRange.Characters.count
                        If IsColorInRange(textRange.Characters(i).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                           textRange.Characters(i).Font.color.RGB <> ignoreColor Then
                            textHasSearchColor = True
                            Exit For
                        End If
                    Next i
                    
                    ' Fr�ga f�r hela textrutan om f�rgen hittas
                    If textHasSearchColor Then
                        slide.Select
                        shape.Select
                        
                        userResponse = MsgBox("Svart text hittad p� Slide " & slide.SlideIndex & "." & vbCrLf & _
                        "Vill du byta textf�rg till bl�tt?", vbYesNoCancel + vbQuestion, "Bekr�fta f�rg�ndring")
                        
                        If userResponse = vbYes Then
                            For i = 1 To textRange.Characters.count
                                If IsColorInRange(textRange.Characters(i).Font.color.RGB, searchColorLowerBound, searchColorUpperBound) And _
                                   textRange.Characters(i).Font.color.RGB <> ignoreColor Then
                                    textRange.Characters(i).Font.color.RGB = replaceColor
                                End If
                            Next i
                            TotalCount = TotalCount + 1
                        ElseIf userResponse = vbCancel Then
                            MsgBox "Makrot avbr�ts av anv�ndaren.", vbExclamation, "Avbrutet"
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
                
                ' Kontrollera om n�gon cell har s�kf�rgen (och ignorera vit text)
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
                
                ' Fr�ga f�r hela tabellen om f�rgen hittas
                If textHasSearchColor Then
                    slide.Select
                    shape.Select
                    
                    userResponse = MsgBox("Svart text hittad i en tabell p� Slide " & slide.SlideIndex & "." & vbCrLf & _
                    "Vill du byta svart text till bl�tt?", vbYesNoCancel + vbQuestion, "Bekr�fta f�rg�ndring")
                    
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
                        MsgBox "Makrot avbr�ts av anv�ndaren.", vbExclamation, "Avbrutet"
                        Exit Sub
                    End If
                End If
            End If
        Next shape
    Next slide
    
    On Error GoTo 0 ' �terst�ll normal felhantering
    
    ' Bekr�ftelsemeddelande
    MsgBox "Du bytte f�rg p� " & TotalCount & " textrutor, tabeller och former fr�n svart till bl�tt.", vbInformation, "F�rg�ndringar slutf�rda"
End Sub

' Funktion f�r att kontrollera om en f�rg �r inom det angivna intervallet
Function IsColorInRange(color As Long, lowerBound As Long, upperBound As Long) As Boolean
    If color >= lowerBound And color <= upperBound Then
        IsColorInRange = True
    Else
        IsColorInRange = False
    End If
End Function







