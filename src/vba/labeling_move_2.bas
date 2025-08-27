Attribute VB_Name = "labeling_move_2"
Option Explicit

' Store overlapping label pairs (assuming max 100 pairs)
Public OverlapPairs(1 To 100, 2) As String ' Column 1 = Label 1, Column 2 = Label 2
Public TotalOverlapPairs As Integer ' Renamed from OverlapCount to avoid conflicts

Sub labeling_move_2()
    Dim sld As slide
    Dim chrt As chart
    Dim ser As series
    Dim lbl1 As dataLabel, lbl2 As dataLabel
    Dim i As Integer, j As Integer
    Dim lbl1Left As Double, lbl1Top As Double, lbl1Right As Double, lbl1Bottom As Double
    Dim lbl2Left As Double, lbl2Top As Double, lbl2Right As Double, lbl2Bottom As Double

    ' Reset stored overlaps
    TotalOverlapPairs = 0
    Erase OverlapPairs

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

    ' Ensure the series has data labels
    If ser.DataLabels.count < 2 Then
        MsgBox "Not enough data labels to check for overlaps.", vbExclamation
        Exit Sub
    End If

    ' Immediate window print
    Debug.Print "---------------------------------------------"
    Debug.Print "Identifying Overlapping Pairs in Chart..."
    Debug.Print "---------------------------------------------"

    ' Loop through all label pairs
    For i = 1 To ser.Points.count - 1
        If ser.Points(i).HasDataLabel Then
            Set lbl1 = ser.Points(i).dataLabel
            
            ' Print label 1 info
            Debug.Print "Checking Label " & i & ": [" & lbl1.text & "]"

            ' Get label 1 dimensions
            If Not GetLabelBounds(lbl1, lbl1Left, lbl1Top, lbl1Right, lbl1Bottom) Then GoTo SkipLabel1
            
            For j = i + 1 To ser.Points.count
                If ser.Points(j).HasDataLabel Then
                    Set lbl2 = ser.Points(j).dataLabel
                    
                    ' Print label 2 info
                    Debug.Print "Comparing to Label " & j & ": [" & lbl2.text & "]"

                    ' Get label 2 dimensions
                    If Not GetLabelBounds(lbl2, lbl2Left, lbl2Top, lbl2Right, lbl2Bottom) Then GoTo SkipLabel2
                    
                    ' **Check for overlap**
                    If (lbl1Left < lbl2Right And lbl1Right > lbl2Left) And _
                       (lbl1Top < lbl2Bottom And lbl1Bottom > lbl2Top) Then

                        ' Increase overlap counter
                        TotalOverlapPairs = TotalOverlapPairs + 1

                        ' Store the overlapping pair
                        If TotalOverlapPairs <= 100 Then
                            OverlapPairs(TotalOverlapPairs, 1) = lbl1.text
                            OverlapPairs(TotalOverlapPairs, 2) = lbl2.text
                            
                            ' Print to Immediate Window
                            Debug.Print "Overlap Found #" & TotalOverlapPairs & ": " & _
                                        "[" & lbl1.text & "] overlaps with [" & lbl2.text & "]"
                        End If
                    End If
                End If
SkipLabel2:
            Next j
        End If
SkipLabel1:
    Next i

    ' Print total count
    Debug.Print "---------------------------------------------"
    Debug.Print "Total overlapping pairs found: " & TotalOverlapPairs
    Debug.Print "---------------------------------------------"

    ' Print all stored overlapping pairs in the Immediate Window
    Debug.Print "Stored Overlapping Pairs (Public Variable):"
    If TotalOverlapPairs > 0 Then
        For i = 1 To TotalOverlapPairs
            Debug.Print "Variable OverlapPairs(" & i & ", 1) = [" & OverlapPairs(i, 1) & "]"
            Debug.Print "Variable OverlapPairs(" & i & ", 2) = [" & OverlapPairs(i, 2) & "]"
        Next i
    Else
        Debug.Print "No overlaps stored."
    End If

    ' Print all functions used
    Debug.Print "---------------------------------------------"
    Debug.Print "List of Functions Used in This Module:"
    Debug.Print "- IdentifyOverlappingPairs"
    Debug.Print "- GetLabelBounds"
    Debug.Print "- IsNaN"
    Debug.Print "---------------------------------------------"
End Sub

' Function to get label bounds, avoiding errors
Function GetLabelBounds(lbl As dataLabel, leftVal As Double, topVal As Double, rightVal As Double, bottomVal As Double) As Boolean
    On Error Resume Next
    leftVal = lbl.left
    topVal = lbl.Top
    rightVal = leftVal + lbl.width
    bottomVal = topVal + lbl.height
    If Err.Number <> 0 Or IsNaN(leftVal) Or IsNaN(topVal) Then
        Err.Clear
        GetLabelBounds = False
    Else
        GetLabelBounds = True
    End If
    On Error GoTo 0
End Function

' Function to check for NaN values
Function IsNaN(value As Double) As Boolean
    IsNaN = (value <> value)
End Function


