Attribute VB_Name = "LD_1_3"
Sub CopyChartFromSlideWithTitle()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa ett tomt LIE DETECTOR-diagram?", vbYesNo + vbQuestion, "Bekräfta")
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

