Attribute VB_Name = "SS_NEW_CHART_1"
Sub CopyChartFromSlideWithTitle2()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa ett tomt SUBSTANSMATRIS-diagram?", vbYesNo + vbQuestion, "Confirm")
    If response = vbNo Then Exit Sub

    Dim sourceSlide As slide
    Dim destSlide As slide
    Dim sourceChart As shape
    Dim shp As shape
    Dim foundSlide As Boolean
    foundSlide = False

    ' Hitta sliden med rubriken "Diagram 2"
    For Each sourceSlide In ActivePresentation.Slides
        For Each shp In sourceSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Trim(shp.TextFrame.textRange.text) = "Diagram 2" Then
                        foundSlide = True
                        Exit For
                    End If
                End If
            End If
        Next shp
        If foundSlide Then Exit For
    Next sourceSlide

    If Not foundSlide Then Exit Sub

    ' Leta upp det enda diagrammet (msoChart) på sliden
    Set sourceChart = Nothing
    Dim chartCount As Integer: chartCount = 0
    For Each shp In sourceSlide.Shapes
        If shp.Type = msoChart Then
            Set sourceChart = shp
            chartCount = chartCount + 1
        End If
    Next shp

    If chartCount = 0 Then Exit Sub
    If chartCount > 1 Then
        MsgBox "More than one chart found on 'Diagram 2' slide. Please ensure there's only one.", vbExclamation
        Exit Sub
    End If

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

