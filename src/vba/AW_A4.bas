Attribute VB_Name = "AW_A4"
Sub CopyChartFromSlideWithTitle3()
    Dim response As VbMsgBoxResult
    response = MsgBox("Vill du skapa ett tomt AWARENESS-diagram?", vbYesNo + vbQuestion, "Bekräfta")
    If response = vbNo Then Exit Sub

    Dim sourceSlide As slide
    Dim destSlide As slide
    Dim sourceChart As shape
    Dim shp As shape
    Dim foundSlide As Boolean
    foundSlide = False

    ' Hitta sliden med rubriken "Diagram 3"
    For Each sourceSlide In ActivePresentation.Slides
        For Each shp In sourceSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Trim(shp.TextFrame.textRange.text) = "Diagram 3" Then
                        foundSlide = True
                        Exit For
                    End If
                End If
            End If
        Next shp
        If foundSlide Then Exit For
    Next sourceSlide

    If Not foundSlide Then Exit Sub

    ' Hämta första diagrammet på sliden, oavsett namn
    Set sourceChart = Nothing
    For Each shp In sourceSlide.Shapes
        If shp.Type = msoChart Then
            Set sourceChart = shp
            Exit For
        End If
    Next shp

    If sourceChart Is Nothing Then Exit Sub

    ' Klistra in på den aktiva sliden
    sourceChart.Copy
    Set destSlide = ActiveWindow.View.slide
    Dim copiedChart As shape
    Set copiedChart = destSlide.Shapes.Paste(1)

    ' Placera kopian på samma plats och storlek
    With copiedChart
        .left = sourceChart.left
        .Top = sourceChart.Top
        .width = sourceChart.width
        .height = sourceChart.height
    End With

    MsgBox "Nytt Awareness-diagram inkopierat (obs osynligt - klicka 'Get Data' ", vbInformation
End Sub

