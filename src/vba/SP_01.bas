Attribute VB_Name = "SP_01"
Sub SP_01()
    Dim pptSlide As slide
    Dim rightmostTable As shape
    Dim rightMostPosition As Single
    Dim table As shape

    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Initialize variables to track the rightmost table
    Set rightmostTable = Nothing
    rightMostPosition = 0

    ' Loop through all shapes on the slide to find the rightmost table
    For Each table In pptSlide.Shapes
        If table.Type = msoTable Then
            If table.left + table.width > rightMostPosition Then
                Set rightmostTable = table
                rightMostPosition = table.left + table.width
            End If
        End If
    Next table

    ' If a rightmost table is found, rename it as "TARGET"
    If Not rightmostTable Is Nothing Then
        rightmostTable.Name = "TARGET"
    Else
        MsgBox "No table found on the slide.", vbExclamation
    End If

End Sub

