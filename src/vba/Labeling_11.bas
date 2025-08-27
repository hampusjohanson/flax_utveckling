Attribute VB_Name = "Labeling_11"
Option Explicit

Sub AdjustFlankLabels1()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer, j As Integer
    Dim lblLeft() As Double, lblTop() As Double
    Dim lblText() As String
    Dim isLeftFlank() As Boolean, isRightFlank() As Boolean, isTopFlank() As Boolean, isBottomFlank() As Boolean
    Dim numPoints As Integer
    Dim leftmostIndex As Integer, rightmostIndex As Integer, topmostIndex As Integer, bottommostIndex As Integer
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    
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
                ReDim isRightFlank(1 To numPoints)
                ReDim isTopFlank(1 To numPoints)
                ReDim isBottomFlank(1 To numPoints)
                
                minX = 1E+30: maxX = -1E+30
                minY = 1E+30: maxY = -1E+30
                leftmostIndex = -1: rightmostIndex = -1
                topmostIndex = -1: bottommostIndex = -1
                
                ' Store label positions
                For i = 1 To numPoints
                    If DataPointScreenCoords(srs, i, chrt, lblLeft(i), lblTop(i)) Then
                        lblText(i) = srs.Points(i).dataLabel.text
                        isLeftFlank(i) = True
                        isRightFlank(i) = True
                        isTopFlank(i) = True
                        isBottomFlank(i) = True
                        
                        If lblLeft(i) < minX Then minX = lblLeft(i): leftmostIndex = i
                        If lblLeft(i) > maxX Then maxX = lblLeft(i): rightmostIndex = i
                        If lblTop(i) < minY Then minY = lblTop(i): topmostIndex = i
                        If lblTop(i) > maxY Then maxY = lblTop(i): bottommostIndex = i
                    End If
                Next i
                
                ' Check for flank labels
                For i = 1 To numPoints
                    For j = 1 To numPoints
                        If i <> j Then
                            If lblLeft(j) < lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < 15 Then
                                isLeftFlank(i) = False
                            End If
                            If lblLeft(j) > lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < 15 Then
                                isRightFlank(i) = False
                            End If
                            If lblTop(j) < lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < 15 Then
                                isTopFlank(i) = False
                            End If
                            If lblTop(j) > lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < 15 Then
                                isBottomFlank(i) = False
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels
                Debug.Print "Adjusted Flank Labels:"
                For i = 1 To numPoints
                    If isLeftFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (Left Flank)"
                        If i = leftmostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionLeft
                        End If
                    End If
                    If isRightFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (Right Flank)"
                        If i = rightmostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionRight
                        End If
                    End If
                    If isTopFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (Top Flank)"
                        srs.Points(i).dataLabel.Position = xlLabelPositionAbove
                    End If
                    If isBottomFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (Bottom Flank)"
                        srs.Points(i).dataLabel.Position = xlLabelPositionBelow
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

