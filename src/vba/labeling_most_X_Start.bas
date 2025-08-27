Attribute VB_Name = "labeling_most_X_Start"
Option Explicit

' Global variables
Public LabelToMove As dataLabel
Public LabelToMoveText As String
Public OriginalColor As Long
Public topLabel As dataLabel
Public checkedPairs() As Boolean
Public totalLabels As Integer
Public labelTexts() As String
Public SkippedLabels() As String
Public SkippedLabelCount As Integer

Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = (Not IsError(LBound(arr, 1))) And (Not IsError(UBound(arr, 1)))
    On Error GoTo 0
End Function

Sub labeling_find_and_ask_multiple()
    Dim sld As slide
    Dim chrt As chart
    Dim ser As series
    Dim lbls() As dataLabel
    Dim labelLefts() As Double, labelRights() As Double
    Dim labelTops() As Double, labelBottoms() As Double
    Dim i As Integer, j As Integer
    Dim overlapWidth As Double, overlapHeight As Double, overlapArea As Double
    Dim maxOverlapArea As Double
    Dim isSkipped As Boolean

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

    ' Store label data in arrays for faster access
    totalLabels = ser.Points.count
    ReDim lbls(1 To totalLabels)
    ReDim labelTexts(1 To totalLabels)
    ReDim labelLefts(1 To totalLabels)
    ReDim labelRights(1 To totalLabels)
    ReDim labelTops(1 To totalLabels)
    ReDim labelBottoms(1 To totalLabels)

    ' Ensure checkedPairs is allocated
    If Not IsArrayAllocated(checkedPairs) Then ReDim checkedPairs(1 To totalLabels, 1 To totalLabels)

    ' Ensure skipped labels array is allocated
    If SkippedLabelCount = 0 Then ReDim SkippedLabels(1 To totalLabels)

    ' Fill arrays with label positions
    For i = 1 To totalLabels
        If ser.Points(i).HasDataLabel Then
            Set lbls(i) = ser.Points(i).dataLabel
            labelTexts(i) = lbls(i).text
            labelLefts(i) = lbls(i).left
            labelRights(i) = labelLefts(i) + lbls(i).width
            labelTops(i) = lbls(i).Top
            labelBottoms(i) = labelTops(i) + lbls(i).height
        Else
            labelTexts(i) = "" ' No label
        End If
    Next i

    ' Find the most overlapping label **that is NOT in the skipped list**
    maxOverlapArea = 0
    Set topLabel = Nothing

    For i = 1 To totalLabels - 1
        If labelTexts(i) <> "" Then
            For j = i + 1 To totalLabels
                If labelTexts(j) <> "" And Not checkedPairs(i, j) Then

                    ' **Check if label is in the skipped list**
                    isSkipped = False
                    Dim k As Integer
                    For k = 1 To SkippedLabelCount
                        If SkippedLabels(k) = labelTexts(i) Or SkippedLabels(k) = labelTexts(j) Then
                            isSkipped = True
                            Exit For
                        End If
                    Next k

                    ' **Skip this pair if it's been skipped before**
                    If isSkipped Then GoTo SkipThisPair

                    ' Check for overlap
                    If (labelLefts(i) < labelRights(j) And labelRights(i) > labelLefts(j)) And _
                       (labelTops(i) < labelBottoms(j) And labelBottoms(i) > labelTops(j)) Then

                        ' Calculate overlap area
                        overlapWidth = Min(labelRights(i), labelRights(j)) - Max(labelLefts(i), labelLefts(j))
                        overlapHeight = Min(labelBottoms(i), labelBottoms(j)) - Max(labelTops(i), labelTops(j))
                        overlapArea = overlapWidth * overlapHeight

                        ' Check if this is the largest overlap
                        If overlapArea > maxOverlapArea Then
                            maxOverlapArea = overlapArea

                            ' Determine the top-most label
                            If labelTops(i) < labelTops(j) Then
                                Set topLabel = lbls(i)
                            Else
                                Set topLabel = lbls(j)
                            End If
                        End If
                    End If
                End If
SkipThisPair:
            Next j
        End If
    Next i

    ' No overlaps found
    If maxOverlapArea = 0 Then
        MsgBox "No more overlapping labels found.", vbExclamation
        Exit Sub
    End If

    ' Ensure a valid label is selected before proceeding
    If topLabel Is Nothing Then
        MsgBox "Error: No overlapping label found.", vbExclamation
        Exit Sub
    End If

    ' Store the selected label globally
    OriginalColor = topLabel.Font.color
    topLabel.Font.color = RGB(255, 0, 0) ' Highlight in red

    ' Use "Set" for object assignment
    Set LabelToMove = topLabel
    LabelToMoveText = topLabel.text ' Backup

    ' Reset & Show UserForm with only the selected label
    Unload frmLabelOverlap
    DoEvents ' Refresh UI before reopening
    frmLabelOverlap.lblOverlapInfo.Caption = "Do you want to move '" & LabelToMoveText & "'?"
    frmLabelOverlap.Show vbModeless
End Sub

' Helper functions
Function Min(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min = a Else Min = b
End Function

Function Max(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max = a Else Max = b
End Function


