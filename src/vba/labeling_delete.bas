Attribute VB_Name = "labeling_delete"

Sub DeleteAllDataLabels()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim i As Integer

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
                
                ' Remove all data labels
                srs.HasDataLabels = False
                
                ' Debug info
                Debug.Print "Deleted all data labels."
            End If
        End If
    Next shp
End Sub
