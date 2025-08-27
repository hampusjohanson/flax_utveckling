Attribute VB_Name = "Lines_Larger_Right"
Sub Adjust_Right_Chart_By_Axis_Settings()

    Dim pptSlide As slide
    Dim leftChartShape As shape
    Dim rightChartShape As shape
    Dim leftChart As chart
    Dim rightChart As chart
    Dim leftAxis As Axis
    Dim rightAxis As Axis
    Dim leftVisibleCount As Long
    Dim rightVisibleCount As Long
    
    Set pptSlide = ActiveWindow.View.slide
    
    ' Hämta diagrammen
    On Error Resume Next
    Set leftChartShape = pptSlide.Shapes("left_chart")
    Set rightChartShape = pptSlide.Shapes("right_chart")
    On Error GoTo 0
    
    If leftChartShape Is Nothing Or rightChartShape Is Nothing Then
        Debug.Print "Chart shapes not found."
        Exit Sub
    End If
    
    If Not leftChartShape.hasChart Or Not rightChartShape.hasChart Then
        Debug.Print "One or both shapes are not charts."
        Exit Sub
    End If
    
    Set leftChart = leftChartShape.chart
    Set rightChart = rightChartShape.chart
    
    ' Hämta Y-axlarna (kategori-axeln)
    Set leftAxis = leftChart.Axes(xlValue, xlPrimary) ' För scatter är Y oftast Value axis
    Set rightAxis = rightChart.Axes(xlValue, xlPrimary)

    ' Räkna synliga "rader" baserat på axelns bounds
    leftVisibleCount = leftAxis.MaximumScale - leftAxis.MinimumScale + 1
    rightVisibleCount = rightAxis.MaximumScale - rightAxis.MinimumScale + 1
    
    Debug.Print "Visible rows - Left: " & leftVisibleCount & ", Right: " & rightVisibleCount
    
    If rightVisibleCount = leftVisibleCount + 1 Then
        With rightChartShape
            Debug.Print "Adjusting right_chart height."
            .height = 11.58 * 28.35 ' cm till points
        End With
    Else
        Debug.Print "No adjustment needed."
    End If

End Sub

