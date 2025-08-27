Attribute VB_Name = "AA_Series_test_7"
Sub AA_Series_7()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObj As chart
    Dim seriesCount As Integer
    Dim numCategories As Integer
    Dim i As Integer, j As Integer
    Dim visibleSeries() As series
    Dim visibleSeriesCount As Integer
    Dim sumValues() As Double
    Dim topSeries As series

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the chart named "Awareness"
    Set chartShape = Nothing
    For Each chartShape In pptSlide.Shapes
        If chartShape.Name = "Awareness" Then
            Set chartObj = chartShape.chart
            Exit For
        End If
    Next chartShape

    ' Check if the chart is found
    If chartShape Is Nothing Then
        MsgBox "No chart named 'Awareness' found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Get total number of series
    seriesCount = chartObj.SeriesCollection.count
    Debug.Print "Total series count: " & seriesCount

    ' Identify all visible series
    ReDim visibleSeries(1 To seriesCount)
    visibleSeriesCount = 0
    For j = 1 To seriesCount
        If chartObj.SeriesCollection(j).Format.Fill.visible Then
            visibleSeriesCount = visibleSeriesCount + 1
            Set visibleSeries(visibleSeriesCount) = chartObj.SeriesCollection(j)
        End If
    Next j

    ' If fewer than 3 visible series are found
    If visibleSeriesCount < 3 Then
        MsgBox "Less than three visible series found.", vbExclamation
        Exit Sub
    End If

    ' Determine the number of categories (assumes all series have the same number of points)
    numCategories = visibleSeries(1).Points.count
    Debug.Print "Categories (X-axis count): " & numCategories

    ' Ensure there's at least one category
    If numCategories = 0 Then
        MsgBox "No categories found in the chart.", vbExclamation
        Exit Sub
    End If

    ' Get the topmost series (last visible one)
    Set topSeries = visibleSeries(visibleSeriesCount)

    ' Initialize array for storing sum values
    ReDim sumValues(1 To numCategories)

    ' Loop through each category to calculate the sum of the **first two** visible series
    For i = 1 To numCategories
        sumValues(i) = visibleSeries(1).values(i) + visibleSeries(2).values(i)

        Debug.Print "Category " & i & " - First Two Visible Values: " & _
                    Format(visibleSeries(1).values(i), "0.00%") & " + " & Format(visibleSeries(2).values(i), "0.00%") & " = " & _
                    Format(sumValues(i), "0.00%")

        ' Apply as data label to the top series
        topSeries.Points(i).dataLabel.text = Format(sumValues(i), "0%")
    Next i

    ' Set data label font color and make it bold
    With topSeries.DataLabels
        .Font.color = RGB(17, 21, 66) ' Dark blue font color
        .Font.Bold = True ' Make text bold
    End With

    Debug.Print "? AA_Series_7 executed correctly!"
End Sub


' Sorting function to arrange values in ascending order
Sub BubbleSortAscending(arr As Variant)
    Dim i As Integer, j As Integer
    Dim temp As Double
    Dim n As Integer
    n = UBound(arr)
    
    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                ' Swap
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub



