Attribute VB_Name = "CAP_NEW"
Sub CAP_NEW()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub
Sub SP_NEW()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle5"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub



Sub CopyChartFromSlideWithTitle4()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa ett tomt CAPITALIZATION-diagram?", vbYesNo + vbQuestion, "Confirm")
    If response = vbNo Then Exit Sub

    Dim sourceSlide As slide
    Dim destSlide As slide
    Dim sourceChart As shape
    Dim shp As shape
    Dim foundSlide As Boolean
    foundSlide = False

    ' Hitta sliden med rubriken "Diagram 1"
    For Each sourceSlide In ActivePresentation.Slides
        For Each shp In sourceSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Trim(shp.TextFrame.textRange.text) = "Diagram 1" Then
                        foundSlide = True
                        Exit For
                    End If
                End If
            End If
        Next shp
        If foundSlide Then Exit For
    Next sourceSlide

    If Not foundSlide Then Exit Sub

    ' Hämta diagrammet med namnet "Chart_Type_1"
    Set sourceChart = Nothing
    For Each shp In sourceSlide.Shapes
        If shp.Name = "Chart_Type_1" Then
            Set sourceChart = shp
            Exit For
        End If
    Next shp

    If sourceChart Is Nothing Then Exit Sub

    ' Kopiera och klistra in på den aktiva sliden
    sourceChart.Copy
    Set destSlide = ActiveWindow.View.slide
    Dim copiedChart As shape
    Set copiedChart = destSlide.Shapes.Paste(1)

    ' Placera på samma position och storlek
    With copiedChart
        .left = sourceChart.left
        .Top = sourceChart.Top
        .width = sourceChart.width
        .height = sourceChart.height
    End With

End Sub




Sub CopyChartFromSlideWithTitle5()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa ett tomt SALES PREMIUM-diagram?", vbYesNo + vbQuestion, "Confirm")
    If response = vbNo Then Exit Sub

    Dim sourceSlide As slide
    Dim destSlide As slide
    Dim sourceChart As shape
    Dim shp As shape
    Dim foundSlide As Boolean
    foundSlide = False

    ' Hitta sliden med rubriken "Diagram 1"
    For Each sourceSlide In ActivePresentation.Slides
        For Each shp In sourceSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Trim(shp.TextFrame.textRange.text) = "Diagram 1" Then
                        foundSlide = True
                        Exit For
                    End If
                End If
            End If
        Next shp
        If foundSlide Then Exit For
    Next sourceSlide

    If Not foundSlide Then Exit Sub

    ' Hämta diagrammet med namnet "Chart_Type_1"
    Set sourceChart = Nothing
    For Each shp In sourceSlide.Shapes
        If shp.Name = "Chart_Type_1" Then
            Set sourceChart = shp
            Exit For
        End If
    Next shp

    If sourceChart Is Nothing Then Exit Sub

    ' Kopiera och klistra in på den aktiva sliden
    sourceChart.Copy
    Set destSlide = ActiveWindow.View.slide
    Dim copiedChart As shape
    Set copiedChart = destSlide.Shapes.Paste(1)

    ' Placera på samma position och storlek
    With copiedChart
        .left = sourceChart.left
        .Top = sourceChart.Top
        .width = sourceChart.width
        .height = sourceChart.height
    End With

End Sub



