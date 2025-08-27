Attribute VB_Name = "labeling_4"
Option Explicit

Sub IdentifyAndMoveTopFlankLabels()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer, j As Integer
    Dim lblLeft() As Double, lblTop() As Double
    Dim lblText() As String
    Dim isTopFlank() As Boolean
    Dim numPoints As Integer
    Dim topmostIndex As Integer
    Dim minY As Double
    
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
                ReDim isTopFlank(1 To numPoints)
                minY = 1E+30
                topmostIndex = -1
                
                ' Store label positions
                For i = 1 To numPoints
                    If DataPointScreenCoords(srs, i, chrt, lblLeft(i), lblTop(i)) Then
                        lblText(i) = srs.Points(i).dataLabel.text
                        isTopFlank(i) = True ' Assume it's a top-flank label initially
                        If lblTop(i) < minY Then
                            minY = lblTop(i)
                            topmostIndex = i
                        End If
                    End If
                Next i
                
                ' Check which points have another point further up in their horizontal range
                For i = 1 To numPoints
                    For j = 1 To numPoints
                        If i <> j Then
                            If lblTop(j) < lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < 15 Then ' Check upward & horizontal overlap
                                isTopFlank(i) = False ' This point has another above it
                                Exit For
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels
                Debug.Print "Top-Flank Labels (No Label Further Up):"
                For i = 1 To numPoints
                    If isTopFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (X: " & lblLeft(i) & ", Y: " & lblTop(i) & ")"
                        If i = topmostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionBelow ' Move topmost label below
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove ' Move others above
                        End If
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

Sub IdentifyAndMoveBottomFlankLabels()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer, j As Integer
    Dim lblLeft() As Double, lblTop() As Double
    Dim lblText() As String
    Dim isBottomFlank() As Boolean
    Dim numPoints As Integer
    Dim bottommostIndex As Integer
    Dim maxY As Double
    
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
                ReDim isBottomFlank(1 To numPoints)
                maxY = -1E+30
                bottommostIndex = -1
                
                ' Store label positions
                For i = 1 To numPoints
                    If DataPointScreenCoords(srs, i, chrt, lblLeft(i), lblTop(i)) Then
                        lblText(i) = srs.Points(i).dataLabel.text
                        isBottomFlank(i) = True ' Assume it's a bottom-flank label initially
                        If lblTop(i) > maxY Then
                            maxY = lblTop(i)
                            bottommostIndex = i
                        End If
                    End If
                Next i
                
                ' Check which points have another point further down in their horizontal range
                For i = 1 To numPoints
                    For j = 1 To numPoints
                        If i <> j Then
                            If lblTop(j) > lblTop(i) And Abs(lblLeft(j) - lblLeft(i)) < 15 Then ' Check downward & horizontal overlap
                                isBottomFlank(i) = False ' This point has another below it
                                Exit For
                            End If
                        End If
                    Next j
                Next i
                
                ' Move labels
                Debug.Print "Bottom-Flank Labels (No Label Further Down):"
                For i = 1 To numPoints
                    If isBottomFlank(i) And lblText(i) <> "False" Then
                        Debug.Print " - " & lblText(i) & " (X: " & lblLeft(i) & ", Y: " & lblTop(i) & ")"
                        If i = bottommostIndex Then
                            srs.Points(i).dataLabel.Position = xlLabelPositionAbove ' Move bottommost label above
                        Else
                            srs.Points(i).dataLabel.Position = xlLabelPositionBelow ' Move others below
                        End If
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

