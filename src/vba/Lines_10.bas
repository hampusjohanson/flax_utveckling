Attribute VB_Name = "Lines_10"
Sub Lines_10()
    Dim pptSlide As slide
    Dim chart As chart
    Dim series As series
    Dim i As Integer
    Dim shape As shape
    Dim table As table
    Dim seriesNames As Collection
    Dim rowIndex As Integer
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim targetRow As Integer
    Dim cellShape As shape
    Dim headerRow As Integer
    Dim targetRows As Variant
    Dim j As Integer

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Find the first chart on the slide
    For Each shape In pptSlide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart
            Exit For
        End If
    Next shape

    ' If no chart found, show an error and exit
    If chart Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Create a collection for series names
    Set seriesNames = New Collection

    ' Count how many series should be included (excluding "FALSE" and "FALSKT")
    Dim validSeriesCount As Integer
    validSeriesCount = 0

    For i = 1 To chart.SeriesCollection.count - 1 ' Exclude last series
        Set series = chart.SeriesCollection(i)
        If LCase(series.Name) <> "false" And LCase(series.Name) <> "falskt" Then
            validSeriesCount = validSeriesCount + 1
            seriesNames.Add series.Name
        End If
    Next i

    ' If no valid series remain, show a message and exit
    If validSeriesCount = 0 Then
        MsgBox "No valid series found to display in the table.", vbExclamation
        Exit Sub
    End If

    ' Create a table with 2 columns and 'validSeriesCount + 1' rows (for header)
    Set shape = pptSlide.Shapes.AddTable(validSeriesCount + 1, 2)
    Set table = shape.table

    ' Set the name of the table to "Choice1"
    shape.Name = "Choice1"

    ' Set the header for the table
    table.cell(1, 1).shape.TextFrame.textRange.text = "Brand"
    table.cell(1, 2).shape.TextFrame.textRange.text = "Show in Chart (Yes/No)"

    ' Populate the table with valid series names
    rowIndex = 2 ' Start from row 2 as row 1 is the header
    For i = 1 To seriesNames.count
        table.cell(rowIndex, 1).shape.TextFrame.textRange.text = seriesNames(i)
        table.cell(rowIndex, 2).shape.TextFrame.textRange.text = "Yes" ' Default to "Yes"
        rowIndex = rowIndex + 1
    Next i

    ' Resize the table to fit the content
    shape.width = 300
    shape.height = 200
    shape.left = 100
    shape.Top = 100

    ' Apply formatting
    rowCount = validSeriesCount + 1
    colCount = 2
    headerRow = 1

    ' Define the target rows for special formatting
    targetRows = Array(17, 21, 66)

 
            
Dim k As Integer

' Define the target rows for special formatting
targetRows = Array(17, 21, 66)

' Apply special formatting for target rows if they exist
For k = LBound(targetRows) To UBound(targetRows)
    If targetRows(k) <= rowCount And targetRows(k) > 0 Then
        If i = targetRows(k) Then
            cellShape.Fill.ForeColor.RGB = RGB(203, 203, 203) ' Light gray
        End If
    End If
Next k


End Sub

