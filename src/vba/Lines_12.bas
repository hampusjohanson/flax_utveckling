Attribute VB_Name = "Lines_12"
Sub Lines_12()
    Dim pptSlide As slide
    Dim chart As chart
    Dim series As series
    Dim i As Integer
    Dim shape As shape
    Dim tablesOnRight As Collection
    Dim leftMostTable As shape
    Dim rightmostTable As shape
    Dim table As shape

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Initialize the collection for tables on the right side of the chart
    Set tablesOnRight = New Collection

    ' Find the first chart on the slide
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.Type = msoChart Then
            Set chart = shape.chart
            Exit For
        End If
    Next shape
    On Error GoTo 0

    ' If no chart found, show an error and exit
    If chart Is Nothing Then
        MsgBox "No chart found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Loop through all shapes to find tables on the right side of the chart
    For Each shape In pptSlide.Shapes
        If shape.Type = msoTable Then
            ' Check if the table is on the right side of the chart
            If shape.left > chart.Parent.left + chart.Parent.width Then
                tablesOnRight.Add shape
            End If
        End If
    Next shape

    ' If two tables are found on the right side, proceed to rename
    If tablesOnRight.count = 2 Then
        ' Sort tables by their left position (the leftmost one will be Brand_List_1)
        If tablesOnRight(1).left < tablesOnRight(2).left Then
            Set leftMostTable = tablesOnRight(1)
            Set rightmostTable = tablesOnRight(2)
        Else
            Set leftMostTable = tablesOnRight(2)
            Set rightmostTable = tablesOnRight(1)
        End If

        ' Rename the tables
        leftMostTable.Name = "Brand_List_1"
        rightmostTable.Name = "Brand_List_2"

       
    Else
       
    End If
End Sub

