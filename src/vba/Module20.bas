Attribute VB_Name = "Module20"
Sub FindOverlappingLabels()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim lbl1 As dataLabel, lbl2 As dataLabel
    Dim i As Integer, j As Integer
    Dim lbl1Left As Double, lbl1Top As Double, lbl1Width As Double, lbl1Height As Double
    Dim lbl2Left As Double, lbl2Top As Double, lbl2Width As Double, lbl2Height As Double

    ' Get the active slide
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then
        Debug.Print "No active slide detected."
        Exit Sub
    End If

    ' Find the first chart
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set chrt = shp.chart
            If chrt.SeriesCollection.count > 0 Then
                Set srs = chrt.SeriesCollection(1) ' Use first series

                ' Loop through each label
                For i = 1 To srs.Points.count
                    If srs.Points(i).HasDataLabel Then
                        Set lbl1 = srs.Points(i).dataLabel
                        
                        ' Ensure valid values
                        lbl1Left = lbl1.left
                        lbl1Top = lbl1.Top
                        lbl1Width = lbl1.width
                        lbl1Height = lbl1.height

                        ' Handle invalid values
                        If lbl1Width < 1 Then lbl1Width = 20 ' Default width
                        If lbl1Height < 1 Then lbl1Height = 10 ' Default height
                        If lbl1Left = -4 Or lbl1Top = -4 Or lbl1Left > 10000 Or lbl1Top > 10000 Then GoTo NextLabel1

                        ' Compare with other labels
                        For j = i + 1 To srs.Points.count
                            If srs.Points(j).HasDataLabel Then
                                Set lbl2 = srs.Points(j).dataLabel
                                
                                ' Ensure valid values
                                lbl2Left = lbl2.left
                                lbl2Top = lbl2.Top
                                lbl2Width = lbl2.width
                                lbl2Height = lbl2.height

                                ' Handle invalid values
                                If lbl2Width < 1 Then lbl2Width = 20 ' Default width
                                If lbl2Height < 1 Then lbl2Height = 10 ' Default height
                                If lbl2Left = -4 Or lbl2Top = -4 Or lbl2Left > 10000 Or lbl2Top > 10000 Then GoTo NextLabel2

                                ' Check for overlap
                                If (lbl1Left < lbl2Left + lbl2Width And lbl1Left + lbl1Width > lbl2Left) And _
                                   (lbl1Top < lbl2Top + lbl2Height And lbl1Top + lbl1Height > lbl2Top) Then
                                    Debug.Print "Label " & i & " overlaps with Label " & j
                                End If
NextLabel2:
                            End If
                        Next j
NextLabel1:
                    End If
                Next i

            Else
                Debug.Print "No series found in chart."
            End If
            Exit Sub ' Stop after first chart
        End If
    Next shp

    Debug.Print "No chart found on the slide."
End Sub


