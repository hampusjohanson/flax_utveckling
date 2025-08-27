Attribute VB_Name = "Scatter_fix_2b"
' === Modul: Scatter_fix_2b ===

Option Explicit

Sub Scatter_fix_2b()

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim ser As Object
    Dim pt As Object
    Dim i As Integer
    Dim chartFound As Boolean

    Set sld = ActiveWindow.View.slide
    chartFound = False

    ' Leta upp vår kopia
    For Each shp In sld.Shapes
        If shp.Type = msoChart Then
            If shp.Name = "kopia_chart" Then
                Set cht = shp.chart
                chartFound = True
                Exit For
            End If
        End If
    Next shp

    If Not chartFound Then
        Debug.Print "Chart 'kopia_chart' not found on this slide."
        Exit Sub
    End If

    ' Nu jobbar vi bara på kopian
    For Each ser In cht.SeriesCollection
        For i = 1 To ser.Points.count
            Set pt = ser.Points(i)
            On Error Resume Next
            pt.ApplyDataLabels
            pt.dataLabel.Position = xlLabelPositionCenter
            On Error GoTo 0
        Next i
    Next ser

    Debug.Print "Labels set to center on 'kopia_chart'."

End Sub

