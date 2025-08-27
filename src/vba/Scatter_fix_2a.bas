Attribute VB_Name = "Scatter_fix_2a"
' === Modul: Scatter_fix_2a ===

Option Explicit

Sub Scatter_fix_2a()

    Dim sld As slide
    Dim shp As shape
    Dim copiedShape As shape
    Dim chartFound As Boolean

    Set sld = ActiveWindow.View.slide
    chartFound = False

    ' Loopa igenom alla shapes på sliden
    For Each shp In sld.Shapes
        If shp.Type = msoChart Then
            ' Vi duplicerar första hittade diagrammet
            Set copiedShape = shp.Duplicate.Item(1)
            
            ' Flytta kopian långt åt höger
            copiedShape.left = shp.left + 2000
            copiedShape.Top = shp.Top
            
            ' Döp om kopian
            copiedShape.Name = "kopia_chart"
            
            Debug.Print "Duplicated chart as 'kopia_chart'"
            
            chartFound = True
            Exit For ' Vi jobbar bara på första hittade chart
        End If
    Next shp

    If Not chartFound Then
        Debug.Print "No chart found on this slide."
    End If

End Sub

