Attribute VB_Name = "labeling_5"
Option Explicit

Sub AlignDataLabelsLeft()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer
    Dim lbl As dataLabel

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
                
                ' Loop through all data labels
                For i = 1 To srs.Points.count
                    If srs.Points(i).HasDataLabel Then
                        Set lbl = srs.Points(i).dataLabel
                        
                        ' Try to left-align text inside the label
                        On Error Resume Next
                        lbl.HorizontalAlignment = xlHAlignLeft
                        On Error GoTo 0
                        
                        ' Debug info
                        Debug.Print "Aligned label left: " & lbl.text
                    End If
                Next i
            End If
        End If
    Next shp
End Sub

Sub AlignDataLabelsRight()
  
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer
    Dim lbl As dataLabel

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
                
                ' Loop through all data labels
                For i = 1 To srs.Points.count
                    If srs.Points(i).HasDataLabel Then
                        Set lbl = srs.Points(i).dataLabel
                        
                        ' Try to left-align text inside the label
                        On Error Resume Next
                        lbl.HorizontalAlignment = xlHAlignRight
                        On Error GoTo 0
                        
                        ' Debug info
                        Debug.Print "Aligned label right: " & lbl.text
                    End If
                Next i
            End If
        End If
    Next shp
End Sub
