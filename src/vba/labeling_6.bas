Attribute VB_Name = "labeling_6"



Function TrimArray(arr() As String, size As Integer) As String()
    Dim trimmedArr() As String
    ReDim trimmedArr(1 To size)
    Dim i As Integer
    For i = 1 To size
        trimmedArr(i) = arr(i)
    Next i
    TrimArray = trimmedArr
End Function

Function IsInArray(val As String, arr() As String) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


Function EuclideanDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    EuclideanDistance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Function DataPointScreenCoords(srs As series, i As Integer, chrt As chart, xScreen As Double, yScreen As Double) As Boolean
    On Error Resume Next
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim plotLeft As Double, plotTop As Double, plotWidth As Double, plotHeight As Double
    Dim xValue As Double, yValue As Double
    
    xMin = chrt.Axes(xlCategory).MinimumScale
    xMax = chrt.Axes(xlCategory).MaximumScale
    yMin = chrt.Axes(xlValue).MinimumScale
    yMax = chrt.Axes(xlValue).MaximumScale
    
    plotLeft = chrt.PlotArea.left
    plotTop = chrt.PlotArea.Top
    plotWidth = chrt.PlotArea.width
    plotHeight = chrt.PlotArea.height
    
    xValue = srs.xValues(i)
    yValue = srs.values(i)
    
    xScreen = plotLeft + ((xValue - xMin) / (xMax - xMin)) * plotWidth
    yScreen = plotTop + ((yMax - yValue) / (yMax - yMin)) * plotHeight
    
    DataPointScreenCoords = (Err.Number = 0)
    On Error GoTo 0
End Function


