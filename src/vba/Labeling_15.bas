Attribute VB_Name = "Labeling_15"
Option Explicit

Sub AdjustLeftFlankLabelWidth1()
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
    Dim baseWidth As Double: baseWidth = 20 ' Base width in points
    Dim charWidthFactor As Double: charWidthFactor = 5 ' Approximate width per character
    Dim maxWidth As Double: maxWidth = 150 ' Prevent excessive expansion
    
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
                            If lblLeft(j) < lblLeft(i) And Abs(lblTop(j) - lblTop(i)) < 15 Then ' Check leftward & vertical overlap
                                isLeftFlank(i) = False ' This point has another to its left
                                Exit For
                            End If
                        End If
                    Next j
                Next i
                
                ' Adjust label width dynamically based on text length
                Debug.Print "Adjusted Left-Flank Label Widths:"
                For i = 1 To numPoints
                    If isLeftFlank(i) And lblText(i) <> "" And lblText(i) <> "False" Then
                        Dim newWidth As Double
                        newWidth = baseWidth + Len(lblText(i)) * charWidthFactor
                        If newWidth > maxWidth Then newWidth = maxWidth
                        
                        Debug.Print "Expanding " & lblText(i) & " to width: " & newWidth
                        srs.Points(i).dataLabel.width = newWidth ' Expand only horizontally
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

