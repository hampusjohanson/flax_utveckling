Attribute VB_Name = "AA_Series_Test_5"
Sub AA_Series_5()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim newSeries As Object
    Dim numCategories As Integer
    Dim fixedValues() As Variant
    Dim i As Integer

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart named "Awareness"
    For Each chartShape In pptSlide.Shapes
        If chartShape.Name = "Awareness" Then
            Exit For
        End If
    Next chartShape

    ' Check if the chart is found
    If chartShape Is Nothing Then
        MsgBox "No chart named 'Awareness' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Get the number of categories (assumes using X-axis categories)
    numCategories = chartShape.chart.SeriesCollection(1).Points.count

    ' Ensure there's at least one category
    If numCategories = 0 Then
        MsgBox "No categories found in the chart.", vbExclamation
        Exit Sub
    End If

    ' Create an array with 5% values for each category
    ReDim fixedValues(1 To numCategories)
    For i = 1 To numCategories
        fixedValues(i) = 0.08 ' X%
    Next i

    ' Add a new series
    Set newSeries = chartShape.chart.SeriesCollection.newSeries
    newSeries.values = fixedValues

    ' Set the series name
    newSeries.Name = "Fixed X% Series"

    ' Apply percentage formatting to data labels
    newSeries.ApplyDataLabels
    newSeries.DataLabels.NumberFormat = "0%" ' Format as percentage

    ' Set the fill color of the series (#88FFC2)
    With newSeries.Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(136, 255, 194) ' Light Green color
        .Solid
    End With

    ' Set data label font color and make it bold
    With newSeries.DataLabels
        .Font.color = RGB(17, 21, 66) ' Dark Blue font color
        .Font.Bold = True ' Make text bold
    End With

End Sub

