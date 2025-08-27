Attribute VB_Name = "labeling_move_3"
Sub labeling_move_3_identify_leftright()
    Dim sld As slide
    Dim chrt As chart
    Dim ser As series
    Dim lblLeft As dataLabel, lblRight As dataLabel
    Dim lbl1 As dataLabel, lbl2 As dataLabel
    Dim lbl1Left As Double, lbl2Left As Double

    ' Ensure there is at least one overlap
    If TotalOverlapPairs = 0 Then
        MsgBox "No overlapping pairs available.", vbExclamation
        Exit Sub
    End If

    ' Get the current slide
    Set sld = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)

    ' Find the first chart on the slide
    Dim shp As shape
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

    ' Identify the first overlap pair
    Dim label1Text As String, label2Text As String
    label1Text = OverlapPairs(1, 1)
    label2Text = OverlapPairs(1, 2)

    ' Find the corresponding labels in the chart
    Dim i As Integer
    For i = 1 To ser.Points.count
        If ser.Points(i).HasDataLabel Then
            If ser.Points(i).dataLabel.text = label1Text Then
                Set lbl1 = ser.Points(i).dataLabel
            ElseIf ser.Points(i).dataLabel.text = label2Text Then
                Set lbl2 = ser.Points(i).dataLabel
            End If
        End If
    Next i

    ' Ensure both labels were found
    If lbl1 Is Nothing Or lbl2 Is Nothing Then
        MsgBox "One or both labels not found in chart.", vbExclamation
        Exit Sub
    End If

    ' Get Left positions
    lbl1Left = lbl1.left
    lbl2Left = lbl2.left

    ' Determine which is more to the left
    If lbl1Left < lbl2Left Then
        Set lblLeft = lbl1
        Set lblRight = lbl2
    Else
        Set lblLeft = lbl2
        Set lblRight = lbl1
    End If

    ' Print results in Immediate Window
    Debug.Print "---------------------------------------------"
    Debug.Print "First Overlapping Pair Analysis:"
    Debug.Print "Left-most Label: [" & lblLeft.text & "] at Left = " & lblLeft.left
    Debug.Print "Right-most Label: [" & lblRight.text & "] at Left = " & lblRight.left
    Debug.Print "---------------------------------------------"

End Sub

Sub labeling_move_3_identify_topbottom()
    Dim sld As slide
    Dim chrt As chart
    Dim ser As series
    Dim lblTop As dataLabel, lblBottom As dataLabel
    Dim lbl1 As dataLabel, lbl2 As dataLabel
    Dim lbl1Top As Double, lbl2Top As Double

    ' Ensure there is at least one overlap
    If TotalOverlapPairs = 0 Then
        MsgBox "No overlapping pairs available.", vbExclamation
        Exit Sub
    End If

    ' Get the current slide
    Set sld = ActivePresentation.Slides(ActiveWindow.View.slide.SlideIndex)

    ' Find the first chart on the slide
    Dim shp As shape
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

    ' Identify the first overlap pair
    Dim label1Text As String, label2Text As String
    label1Text = OverlapPairs(1, 1)
    label2Text = OverlapPairs(1, 2)

    ' Find the corresponding labels in the chart
    Dim i As Integer
    For i = 1 To ser.Points.count
        If ser.Points(i).HasDataLabel Then
            If ser.Points(i).dataLabel.text = label1Text Then
                Set lbl1 = ser.Points(i).dataLabel
            ElseIf ser.Points(i).dataLabel.text = label2Text Then
                Set lbl2 = ser.Points(i).dataLabel
            End If
        End If
    Next i

    ' Ensure both labels were found
    If lbl1 Is Nothing Or lbl2 Is Nothing Then
        MsgBox "One or both labels not found in chart.", vbExclamation
        Exit Sub
    End If

    ' Get Top positions
    lbl1Top = lbl1.Top
    lbl2Top = lbl2.Top

    ' Determine which is on top and which is on bottom
    If lbl1Top < lbl2Top Then
        Set lblTop = lbl1
        Set lblBottom = lbl2
    Else
        Set lblTop = lbl2
        Set lblBottom = lbl1
    End If

    ' Print results in Immediate Window
    Debug.Print "---------------------------------------------"
    Debug.Print "First Overlapping Pair Analysis:"
    Debug.Print "Top-most Label: [" & lblTop.text & "] at Top = " & lblTop.Top
    Debug.Print "Bottom-most Label: [" & lblBottom.text & "] at Top = " & lblBottom.Top
    Debug.Print "---------------------------------------------"

End Sub


