Attribute VB_Name = "labeling_333"
Option Explicit

Sub IdentifyAndMoveLeftFlankLabels_5()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer, j As Integer
    Dim lblLeft() As Double, lblTop() As Double
    Dim lblText() As String
    Dim isLeftFlank() As Boolean
    Dim numPoints As Integer
    Dim leftmostIndex As Integer
    Dim minX As Double
    
    ' Get the active slide
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then
        Debug.Print "No active slide detected."
        Exit Sub
    End If
    
    ' Loop through shapes to find the first chart
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set chrt = shp.chart
            If chrt.SeriesCollection.count > 0 Then
                Set srs = chrt.SeriesCollection(1) ' First series
                numPoints = srs.Points.count
                
                ' Initialize arrays
                ReDim lblLeft(1 To numPoints)
                ReDim lblTop(1 To numPoints)
                ReDim lblText(1 To numPoints)
                ReDim isLeftFlank(1 To numPoints)
                minX = 1E+30
                leftmostIndex = -1
                
                ' Store label positions
                For i = 1 To numPoints
                    If DataPointScreenCoords(srs, i, chrt, lblLeft(i), lblTop(i)) Then
                        lblText(i) = srs.Points(i).dataLabel.text
                        isLeftFlank(i) = True ' Assume it's a left-flank label initially
                        If lblLeft(i) < minX Then
                            minX = lblLeft(i)
                            leftmostIndex = i
                        End If
                    End If
                Next i
                
                ' Check which points have another point further left in their vertical range
                For i = 1 To numPoints
                    For j = 1 To numPoints
                        If i <> j Then
                            If lblLeft(j) < lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < 5 Then ' Check leftward & vertical overlap
                                isLeftFlank(i) = False ' This point has another to its left
                                Exit For
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels
                Debug.Print "Left-Flank Labels (No Label Further Left):"
                For i = 1 To numPoints
                    If isLeftFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (X: " & lblLeft(i) & ", Y: " & lblTop(i) & ")"
                        If i = leftmostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove ' Move leftmost label above
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionLeft ' Move others to the left
                        End If
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

Sub IdentifyAndMoveRightFlankLabels_5()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer, j As Integer
    Dim lblLeft() As Double, lblTop() As Double
    Dim lblText() As String
    Dim isRightFlank() As Boolean
    Dim numPoints As Integer
    Dim rightmostIndex As Integer
    Dim maxX As Double
    
    ' Get the active slide
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then
        Debug.Print "No active slide detected."
        Exit Sub
    End If
    
    ' Loop through shapes to find the first chart
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set chrt = shp.chart
            If chrt.SeriesCollection.count > 0 Then
                Set srs = chrt.SeriesCollection(1) ' First series
                numPoints = srs.Points.count
                
                ' Initialize arrays
                ReDim lblLeft(1 To numPoints)
                ReDim lblTop(1 To numPoints)
                ReDim lblText(1 To numPoints)
                ReDim isRightFlank(1 To numPoints)
                maxX = -1E+30
                rightmostIndex = -1
                
                ' Store label positions
                For i = 1 To numPoints
                    If DataPointScreenCoords(srs, i, chrt, lblLeft(i), lblTop(i)) Then
                        lblText(i) = srs.Points(i).dataLabel.text
                        isRightFlank(i) = True ' Assume it's a right-flank label initially
                        If lblLeft(i) > maxX Then
                            maxX = lblLeft(i)
                            rightmostIndex = i
                        End If
                    End If
                Next i
                
                ' Check which points have another point further right in their vertical range
                For i = 1 To numPoints
                    For j = 1 To numPoints
                        If i <> j Then
                            If lblLeft(j) > lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < 5 Then ' Check rightward & vertical overlap
                                isRightFlank(i) = False ' This point has another to its right
                                Exit For
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels
                Debug.Print "Right-Flank Labels (No Label Further Right):"
                For i = 1 To numPoints
                    If isRightFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (X: " & lblLeft(i) & ", Y: " & lblTop(i) & ")"
                        If i = rightmostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove ' Move rightmost label above
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionRight ' Move others to the right
                        End If
                    End If
                Next i
            End If
        End If
    Next shp
End Sub






