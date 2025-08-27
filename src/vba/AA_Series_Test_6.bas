Attribute VB_Name = "AA_Series_Test_6"
Sub AA_Series_6()
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

    ' Create an array with 3% values for each category
    ReDim fixedValues(1 To numCategories)
    For i = 1 To numCategories
        fixedValues(i) = 0.05 ' X%
    Next i

    ' Add a new series
    Set newSeries = chartShape.chart.SeriesCollection.newSeries
    newSeries.values = fixedValues

    ' Set the series name
    newSeries.Name = "Fixed % Series"

    ' No fill color
    newSeries.Format.Fill.visible = msoFalse

    ' No data labels
    newSeries.HasDataLabels = False

End Sub

