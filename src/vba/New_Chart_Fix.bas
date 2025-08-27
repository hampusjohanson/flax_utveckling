Attribute VB_Name = "New_Chart_Fix"
Sub UpdateChart()

    Dim pptSlide As slide
    Dim pptChart As chart
    Dim pptShape As shape
    Dim i As Integer
    
    ' Reference the current slide
    Set pptSlide = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)
    Debug.Print "Slide Number: " & pptSlide.SlideIndex ' Immediate Window feedback
    
    ' Loop through all shapes on the slide to find a chart
    For i = 1 To pptSlide.Shapes.count
        Set pptShape = pptSlide.Shapes(i)
        
        ' Check if the shape is a chart
        If pptShape.hasChart Then
            Set pptChart = pptShape.chart
            Debug.Print "Chart found on shape " & i ' Immediate Window feedback
            
            ' Remove gridlines
            pptChart.Axes(xlCategory).MajorGridlines.Delete
            pptChart.Axes(xlValue).MajorGridlines.Delete
            Debug.Print "Gridlines removed!" ' Immediate Window feedback
            
            ' Remove axis titles
            pptChart.Axes(xlCategory).HasTitle = False
            pptChart.Axes(xlValue).HasTitle = False
            Debug.Print "Axis titles removed!" ' Immediate Window feedback
            
            ' Remove axis labels (set to "none")
            pptChart.Axes(xlCategory).TickLabelPosition = xlNone
            pptChart.Axes(xlValue).TickLabelPosition = xlNone
            Debug.Print "Axis labels removed!" ' Immediate Window feedback
            
            ' Resize the chart (11 cm = 311.85 points)
            On Error Resume Next ' In case resizing fails
            pptShape.width = 311.85 ' 11 cm in points
            pptShape.height = 311.85 ' 11 cm in points
            If Err.Number <> 0 Then
                Debug.Print "Error resizing the chart: " & Err.Description
                Err.Clear
            Else
                Debug.Print "Chart resized to 311.85 points x 311.85 points" ' Immediate Window feedback
            End If
            On Error GoTo 0 ' Turn off error handling
            
            ' Set the position of the chart (3.67 cm = 104.75 points, 5.12 cm = 145.23 points)
            pptShape.left = 104.75 ' 3.67 cm in points
            pptShape.Top = 145.23 ' 5.12 cm in points
            Debug.Print "Chart positioned at 104.75 points, 145.23 points" ' Immediate Window feedback
            
            ' Set chart area border (explicitly set the line properties)
            With pptChart.ChartArea.Format.line
                .visible = msoTrue ' Ensure the border is visible
                .Weight = 0.25 ' Set the line weight
                .ForeColor.RGB = RGB(17, 21, 66) ' Set the color to RGB(17, 21, 66)
            End With
            Debug.Print "Chart border set to 0.25 pt, RGB(17, 21, 66)" ' Immediate Window feedback
            
            ' Set axis crossings at 2
            pptChart.Axes(xlCategory).CrossesAt = 0.6
            pptChart.Axes(xlValue).CrossesAt = 0.6
            Debug.Print "Axis crossings set to 2" ' Immediate Window feedback
            
            ' Set font to Arial, color to RGB(17, 21, 66)
            With pptChart.Axes(xlCategory).TickLabels.Font
                .Name = "Arial"
                .size = 10
                .color = RGB(17, 21, 66)
            End With
            With pptChart.Axes(xlValue).TickLabels.Font
                .Name = "Arial"
                .size = 10
                .color = RGB(17, 21, 66)
            End With
            Debug.Print "Axis font set to Arial, RGB(17, 21, 66)" ' Immediate Window feedback
            
            ' Set axis line style to long dash
            With pptChart.Axes(xlCategory).Format.line
                .DashStyle = msoLineDash ' Set to a dashed line style
                .Weight = 0.25 ' Line thickness
                .ForeColor.RGB = RGB(17, 21, 66) ' Line color
            End With
            With pptChart.Axes(xlValue).Format.line
                .DashStyle = msoLineDash ' Set to a dashed line style
                .Weight = 0.25 ' Line thickness
                .ForeColor.RGB = RGB(17, 21, 66) ' Line color
            End With
            Debug.Print "Axis line style set to dashed, RGB(17, 21, 66)" ' Immediate Window feedback
            
            Exit Sub ' Exit the loop after finding the first chart
        End If
        
        ' Output the type of shape if it's not a chart
        Debug.Print "Shape " & i & ": " & pptShape.Type ' Immediate Window feedback
    Next i
    
    Debug.Print "No chart found on slide." ' Immediate Window feedback

End Sub

