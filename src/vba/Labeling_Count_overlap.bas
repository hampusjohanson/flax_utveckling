Attribute VB_Name = "Labeling_Count_overlap"
Option Explicit

Public TotalOverlaps As Integer ' Stores the number of overlaps
Public TotalRealLabels As Integer ' Stores the number of real labels

Sub CountUniqueOverlappingLabels_Debug()
    Dim dbg_sld As slide
    Dim dbg_shp As shape
    Dim dbg_chrt As chart
    Dim dbg_srs As series
    Dim dbg_lblShape1 As dataLabel, dbg_lblShape2 As dataLabel
    Dim dbg_i As Integer, dbg_j As Integer
    Dim dbg_lbl1Left As Double, dbg_lbl1Top As Double, dbg_lbl1Right As Double, dbg_lbl1Bottom As Double
    Dim dbg_lbl2Left As Double, dbg_lbl2Top As Double, dbg_lbl2Right As Double, dbg_lbl2Bottom As Double
    Dim dbg_maxLabels As Integer
    Dim dbg_lbl1Text As String, dbg_lbl2Text As String
    Dim dbg_overlapLabels() As String
    Dim dbg_labelCount As Integer

    ' Reset counters
    TotalOverlaps = 0
    TotalRealLabels = 0

    ' Get the active slide
    Set dbg_sld = ActiveWindow.View.slide
    If dbg_sld Is Nothing Then Exit Sub

    ' Loop through shapes to find the first chart
    For Each dbg_shp In dbg_sld.Shapes
        If dbg_shp.hasChart Then
            Set dbg_chrt = dbg_shp.chart
            If dbg_chrt.SeriesCollection.count > 0 Then
                Set dbg_srs = dbg_chrt.SeriesCollection(1) ' First series
                
                ' Count real data points (ignore "false", "falskt" and empty labels)
                dbg_maxLabels = 0
                For dbg_i = 1 To dbg_srs.Points.count
                    If dbg_srs.Points(dbg_i).HasDataLabel Then
                        dbg_lbl1Text = Trim(LCase(dbg_srs.Points(dbg_i).dataLabel.text))
                        dbg_lbl1Text = Replace(dbg_lbl1Text, Chr(160), "") ' Remove non-breaking spaces
                        
                        ' Debug each label to see what PowerPoint returns
                        Debug.Print "Checking Label " & dbg_i & ": [" & dbg_lbl1Text & "]"

                        If dbg_lbl1Text <> "false" And dbg_lbl1Text <> "falskt" And Len(dbg_lbl1Text) > 0 Then
                            dbg_maxLabels = dbg_maxLabels + 1
                        End If
                    End If
                Next dbg_i
                
                ' Store total real labels
                TotalRealLabels = dbg_maxLabels
                Debug.Print "Total Real Labels in Chart: " & TotalRealLabels
                
                ' Resize array for tracking unique overlapping labels
                ReDim dbg_overlapLabels(1 To dbg_maxLabels) As String
                dbg_labelCount = 0

                ' Loop through all label pairs (i, j) where j > i to avoid duplicate checks
                For dbg_i = 1 To dbg_maxLabels
                    If dbg_srs.Points(dbg_i).HasDataLabel Then
                        Set dbg_lblShape1 = dbg_srs.Points(dbg_i).dataLabel
                        
                        ' Ensure the label is valid before accessing properties
                        If Not dbg_lblShape1 Is Nothing Then
                            On Error Resume Next
                            dbg_lbl1Left = dbg_lblShape1.left
                            dbg_lbl1Top = dbg_lblShape1.Top
                            dbg_lbl1Right = dbg_lbl1Left + dbg_lblShape1.width
                            dbg_lbl1Bottom = dbg_lbl1Top + dbg_lblShape1.height
                            dbg_lbl1Text = dbg_lblShape1.text
                            If Err.Number <> 0 Or IsNaN(dbg_lbl1Left) Or IsNaN(dbg_lbl1Top) Then
                                Err.Clear
                                GoTo SkipLabel1
                            End If
                            On Error GoTo 0
                        Else
                            GoTo SkipLabel1
                        End If

                        ' Check overlaps with other labels
                        For dbg_j = dbg_i + 1 To dbg_maxLabels
                            If dbg_srs.Points(dbg_j).HasDataLabel Then
                                Set dbg_lblShape2 = dbg_srs.Points(dbg_j).dataLabel
                                
                                ' Ensure the label is valid before accessing properties
                                If Not dbg_lblShape2 Is Nothing Then
                                    On Error Resume Next
                                    dbg_lbl2Left = dbg_lblShape2.left
                                    dbg_lbl2Top = dbg_lblShape2.Top
                                    dbg_lbl2Right = dbg_lbl2Left + dbg_lblShape2.width
                                    dbg_lbl2Bottom = dbg_lbl2Top + dbg_lblShape2.height
                                    dbg_lbl2Text = dbg_lblShape2.text
                                    If Err.Number <> 0 Or IsNaN(dbg_lbl2Left) Or IsNaN(dbg_lbl2Top) Then
                                        Err.Clear
                                        GoTo SkipLabel2
                                    End If
                                    On Error GoTo 0
                                Else
                                    GoTo SkipLabel2
                                End If

                                ' **Filter out NaN or extreme values before checking overlap**
                                If Not AreValidBounds(dbg_lbl1Left, dbg_lbl1Right, dbg_lbl1Top, dbg_lbl1Bottom) Or _
                                   Not AreValidBounds(dbg_lbl2Left, dbg_lbl2Right, dbg_lbl2Top, dbg_lbl2Bottom) Then
                                    Debug.Print "Skipping due to invalid bounds."
                                    GoTo SkipLabel2
                                End If

                                ' **Check for overlap**
                                If (dbg_lbl1Left < dbg_lbl2Right And dbg_lbl1Right > dbg_lbl2Left) And _
                                   (dbg_lbl1Top < dbg_lbl2Bottom And dbg_lbl1Bottom > dbg_lbl2Top) Then

                                    ' Increase overlap counter
                                    TotalOverlaps = TotalOverlaps + 1

                                    ' Add label 1 if not already counted
                                    If Not IsInArray_Debug(dbg_lbl1Text, dbg_overlapLabels, dbg_labelCount) Then
                                        dbg_labelCount = dbg_labelCount + 1
                                        dbg_overlapLabels(dbg_labelCount) = dbg_lbl1Text
                                    End If
                                    
                                    ' Add label 2 if not already counted
                                    If Not IsInArray_Debug(dbg_lbl2Text, dbg_overlapLabels, dbg_labelCount) Then
                                        dbg_labelCount = dbg_labelCount + 1
                                        dbg_overlapLabels(dbg_labelCount) = dbg_lbl2Text
                                    End If
                                End If
                            End If
SkipLabel2:
                        Next dbg_j
                    End If
SkipLabel1:
                Next dbg_i
                
                ' Print final count
                Debug.Print "Total Unique Labels Overlapping: " & dbg_labelCount
                Debug.Print "Total Overlapping Pairs Found: " & TotalOverlaps

                ' Print each overlapping label in Immediate Window
                Debug.Print "Unique Overlapping Labels:"
                For dbg_i = 1 To dbg_labelCount
                    Debug.Print "- " & dbg_overlapLabels(dbg_i)
                Next dbg_i
            End If
        End If
    Next dbg_shp
End Sub

' **Helper Function to Check if a Value Exists in an Array**
Function IsInArray_Debug(value As String, arr() As String, arrSize As Integer) As Boolean
    Dim k As Integer
    For k = 1 To arrSize
        If arr(k) = value Then
            IsInArray_Debug = True
            Exit Function
        End If
    Next k
    IsInArray_Debug = False
End Function

' **Helper Function to Filter Invalid Bounds**
Function AreValidBounds(leftVal As Double, rightVal As Double, topVal As Double, bottomVal As Double) As Boolean
    AreValidBounds = (leftVal > -100000 And leftVal < 100000 And _
                      rightVal > -100000 And rightVal < 100000 And _
                      topVal > -100000 And topVal < 100000 And _
                      bottomVal > -100000 And bottomVal < 100000)
End Function

' **Helper Function to Check for NaN Values**
Function IsNaN(value As Double) As Boolean
    IsNaN = (value <> value)
End Function


