Attribute VB_Name = "Lines_NEW"
Sub Lines_NEW()
   
   

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "DeleteChartsWithConfirmation"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CopyChartFromSlideWithTitle_LineChart"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
End Sub



Sub CopyChartFromSlideWithTitle6()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa 2 tomma Linje-diagram?", vbYesNo + vbQuestion, "Confirm")
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


Sub CopyChartFromSlideWithTitle_LineChart()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa 2 tomma Linje-diagram?", vbYesNo + vbQuestion, "Confirm")
    If response = vbNo Then Exit Sub

    Dim sourceSlide As slide
    Dim destSlide As slide
    Dim sourceChartLeft As shape, sourceChartRight As shape
    Dim copiedChart As shape
    Dim shp As shape
    Dim foundSlide As Boolean
    foundSlide = False

    ' Hitta sliden med rubriken "Line chart"
    For Each sourceSlide In ActivePresentation.Slides
        For Each shp In sourceSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Trim(shp.TextFrame.textRange.text) = "Line chart" Then
                        foundSlide = True
                        Exit For
                    End If
                End If
            End If
        Next shp
        If foundSlide Then Exit For
    Next sourceSlide

    If Not foundSlide Then Exit Sub

    ' Hämta diagrammen med namnen "left_chart" och "right_chart"
    Set sourceChartLeft = Nothing
    Set sourceChartRight = Nothing

    For Each shp In sourceSlide.Shapes
        If shp.Name = "left_chart" Then Set sourceChartLeft = shp
        If shp.Name = "right_chart" Then Set sourceChartRight = shp
    Next shp

    If sourceChartLeft Is Nothing And sourceChartRight Is Nothing Then Exit Sub

    Set destSlide = ActiveWindow.View.slide

    ' Kopiera båda om de hittas
    If Not sourceChartLeft Is Nothing Then
        sourceChartLeft.Copy
        Set copiedChart = destSlide.Shapes.Paste(1)
        With copiedChart
            .left = sourceChartLeft.left
            .Top = sourceChartLeft.Top
            .width = sourceChartLeft.width
            .height = sourceChartLeft.height
        End With
    End If

    If Not sourceChartRight Is Nothing Then
        sourceChartRight.Copy
        Set copiedChart = destSlide.Shapes.Paste(1)
        With copiedChart
            .left = sourceChartRight.left
            .Top = sourceChartRight.Top
            .width = sourceChartRight.width
            .height = sourceChartRight.height
        End With
    End If
End Sub


