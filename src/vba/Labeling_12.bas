Attribute VB_Name = "Labeling_12"
Option Explicit

Sub AdjustFlankLabels2()
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
    
    ' === CONTROL PARAMETERS ===
    Dim includeLeft As Boolean: includeLeft = True
    Dim includeRight As Boolean: includeRight = True
    Dim includeTop As Boolean: includeTop = True
    Dim includeBottom As Boolean: includeBottom = True
    Dim verticalThreshold As Double: verticalThreshold = 15 ' Pixels to consider labels as vertically close
    Dim horizontalThreshold As Double: horizontalThreshold = 15 ' Pixels to consider labels as horizontally close
    Dim enableDebug As Boolean: enableDebug = True ' Set to False to suppress debug output
    ' ==========================
    
    ' Get the active slide
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then
        If enableDebug Then Debug.Print "No active slide detected."
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
                            If lblLeft(j) < lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < verticalThreshold Then
                                isLeftFlank(i) = False
                            End If
                            If lblLeft(j) > lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < verticalThreshold Then
                                isRightFlank(i) = False
                            End If
                            If lblTop(j) < lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < horizontalThreshold Then
                                isTopFlank(i) = False
                            End If
                            If lblTop(j) > lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < horizontalThreshold Then
                                isBottomFlank(i) = False
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels with priority order: Left ? Right ? Top ? Bottom
                If enableDebug Then Debug.Print "Adjusted Flank Labels:"
                For i = 1 To numPoints
                    If includeLeft And isLeftFlank(i) And lblText(i) <> "False" Then
                        If enableDebug Then Debug.Print "Moving " & lblText(i) & " to Left"
                        srs.Points(i).dataLabel.Position = xlLabelPositionLeft
                    ElseIf includeRight And isRightFlank(i) And lblText(i) <> "False" Then
                        If enableDebug Then Debug.Print "Moving " & lblText(i) & " to Right"
                        srs.Points(i).dataLabel.Position = xlLabelPositionRight
                    ElseIf includeTop And isTopFlank(i) And lblText(i) <> "False" Then
                        If enableDebug Then Debug.Print "Moving " & lblText(i) & " to Above"
                        srs.Points(i).dataLabel.Position = xlLabelPositionAbove
                    ElseIf includeBottom And isBottomFlank(i) And lblText(i) <> "False" Then
                        If enableDebug Then Debug.Print "Moving " & lblText(i) & " to Below"
                        srs.Points(i).dataLabel.Position = xlLabelPositionBelow
                    End If
                Next i
            End If
        End If
    Next shp
End Sub


