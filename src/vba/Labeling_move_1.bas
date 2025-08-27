Attribute VB_Name = "Labeling_move_1"
Sub Labeling_move_1()
    Dim sld As slide
    Dim chrt As chart
    Dim ser As series
    Dim lbl As dataLabel
    Dim moveX As Single, moveY As Single
    Dim shp As shape
    Dim targetIndex As Integer ' Specific data point number

    ' Set movement values (adjust as needed)
    moveX = -5   ' Move right (+) or left (-)
    moveY = 10   ' Move down (+) or up (-)

    ' Define which data point label to move
    targetIndex = 1  ' Change this to the index of the label you want to move

    ' Get the current slide
    Set sld = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)

    ' Find the first chart on the slide
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set chrt = shp.chart
            Exit For
        End If
    Next shp

    ' Ensure a chart was found
    If chrt Is Nothing Then
        MsgBox "No chart found on this slide.", vbExclamation
        Exit Sub
    End If

    ' Get the first series
    Set ser = chrt.SeriesCollection(1)

    ' Ensure the series has enough data points
    If ser.Points.count < targetIndex Then
        MsgBox "Target index exceeds the number of data points in the series.", vbExclamation
        Exit Sub
    End If

    ' Ensure the selected point has a data label
    If ser.Points(targetIndex).HasDataLabel Then
        Set lbl = ser.Points(targetIndex).dataLabel
        ' Move the label
        lbl.left = lbl.left + moveX
        lbl.Top = lbl.Top + moveY
        Debug.Print "Moved label for data point #" & targetIndex & " by X: " & moveX & ", Y: " & moveY
    Else
        MsgBox "Data point #" & targetIndex & " does not have a label.", vbExclamation
    End If

End Sub

