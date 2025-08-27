Attribute VB_Name = "Lines_Format"
Sub Formatera_linjediagram()
    Dim i As Integer
    For i = 1 To 2 ' Run the main operations twice
        Dim slide As slide
        Dim shape As shape
        Dim leftmostPos As Single
        Dim secondLeftmostPos As Single
        Dim sourceChart As chart, targetChart As chart
        Dim sourceShape As shape
        Dim SeriesIndex As Integer
        Dim pointIndex As Integer

        On Error GoTo ErrorHandler
        Set slide = ActiveWindow.View.slide
        leftmostPos = 9999
        secondLeftmostPos = 9999

        ' Identify leftmost and second leftmost charts
        For Each shape In slide.Shapes
            If shape.hasChart Then
                If shape.left < leftmostPos Then
                    secondLeftmostPos = leftmostPos
                    leftmostPos = shape.left
                    Set targetChart = sourceChart
                    Set sourceShape = shape ' Store the shape containing the source chart
                    Set sourceChart = shape.chart
                ElseIf shape.left < secondLeftmostPos Then
                    secondLeftmostPos = shape.left
                    Set targetChart = shape.chart
                End If
            End If
        Next shape

        ' Check if charts are found
        If sourceChart Is Nothing Or targetChart Is Nothing Then
            MsgBox "Charts not found."
            Exit For
        End If

        ' Copy chart size but not position
        With targetChart.Parent
            .width = sourceChart.Parent.width
            .height = sourceChart.Parent.height
            .Top = sourceChart.Parent.Top ' Align top
        End With

        ' Copy the plot area size and position
        With targetChart.PlotArea
            .width = sourceChart.PlotArea.width
            .height = sourceChart.PlotArea.height
            .left = sourceChart.PlotArea.left
            .Top = sourceChart.PlotArea.Top
        End With

        ' Align chart with specific textbox
        AlignWithTextbox slide, targetChart

        ' Copy data label font properties
        CopyDataLabelFontProperties sourceChart, targetChart

        ' Copy horizontal axis font size
        CopyAxisFontProperties sourceChart, targetChart

        If i = 2 Then ' Show the message only after the second run
            MsgBox "Chart properties, alignment, and data label font properties copied successfully!"
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub CopyDataLabelFontProperties(ByVal sourceChart As chart, ByVal targetChart As chart)
    ' Loop through each series in the source chart
    For SeriesIndex = 1 To sourceChart.SeriesCollection.count
        If sourceChart.SeriesCollection(SeriesIndex).HasDataLabels Then
            ' Check for matching series in target chart
            If SeriesIndex <= targetChart.SeriesCollection.count And _
               targetChart.SeriesCollection(SeriesIndex).HasDataLabels Then

                Dim sourcePointsCount As Integer
                Dim targetPointsCount As Integer

                sourcePointsCount = sourceChart.SeriesCollection(SeriesIndex).Points.count
                targetPointsCount = targetChart.SeriesCollection(SeriesIndex).Points.count

                ' Check if the number of points in the series matches
                If sourcePointsCount = targetPointsCount Then
                    ' Loop through all data points in the series
                    For pointIndex = 1 To sourcePointsCount
                        ' Copy font size, font type, and color from source to target
                        With sourceChart.SeriesCollection(SeriesIndex).Points(pointIndex).dataLabel.Font
                            targetChart.SeriesCollection(SeriesIndex).Points(pointIndex).dataLabel.Font.size = .size
                            targetChart.SeriesCollection(SeriesIndex).Points(pointIndex).dataLabel.Font.Name = .Name
                            targetChart.SeriesCollection(SeriesIndex).Points(pointIndex).dataLabel.Font.color = .color
                        End With
                    Next pointIndex
                Else
                    MsgBox "Mismatch in number of data points between series in charts."
                End If
            Else
                MsgBox "Series " & SeriesIndex & " in target chart is missing data labels."
            End If
        End If
    Next SeriesIndex
End Sub

Sub CopyAxisFontProperties(ByVal sourceChart As chart, ByVal targetChart As chart)
    ' Check if the horizontal axis exists
    If Not sourceChart.Axes(xlCategory) Is Nothing And Not targetChart.Axes(xlCategory) Is Nothing Then
        With sourceChart.Axes(xlCategory).TickLabels.Font
            targetChart.Axes(xlCategory).TickLabels.Font.size = .size
            targetChart.Axes(xlCategory).TickLabels.Font.Name = .Name
            targetChart.Axes(xlCategory).TickLabels.Font.color = .color
        End With
    Else
        MsgBox "One of the charts is missing a horizontal axis."
    End If
End Sub

Sub AlignWithTextbox(ByVal slide As slide, ByVal targetChart As chart)
    Dim shape As shape
    For Each shape In slide.Shapes
        If shape.Type = msoTextBox Then
            If InStr(1, shape.TextFrame.textRange.text, "weaker", vbTextCompare) > 0 Or _
               InStr(1, shape.TextFrame.textRange.text, "svagare", vbTextCompare) > 0 Then
                targetChart.Parent.left = shape.left ' Align left with textbox
                Exit For
            End If
        End If
    Next shape
End Sub


