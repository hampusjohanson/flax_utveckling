Attribute VB_Name = "Scatter_fix_2d"
' === Modul: Scatter_fix_2d ===

Option Explicit

Sub Scatter_fix_2d()

    Dim sld As slide
    Dim shp As shape
    Dim chartDeleted As Boolean

    Set sld = ActiveWindow.View.slide
    chartDeleted = False

    For Each shp In sld.Shapes
        If shp.Type = msoChart Then
            If shp.Name = "kopia_chart" Then
                shp.Delete
                chartDeleted = True
                Exit For
            End If
        End If
    Next shp

    If chartDeleted Then
        Debug.Print "Kopia_chart successfully deleted."
    Else
        Debug.Print "No kopia_chart found on this slide."
    End If

End Sub

