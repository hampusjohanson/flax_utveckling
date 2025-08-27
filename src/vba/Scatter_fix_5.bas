Attribute VB_Name = "Scatter_fix_5"
' === Modul: Scatter_fix_5 ===

Option Explicit

Sub Scatter_fix_5()

    Dim sld As slide
    Dim shp As shape
    Dim chartShape As shape
    Dim tableShape As shape

    Set sld = ActiveWindow.View.slide
    Set chartShape = Nothing
    Set tableShape = Nothing
    
    ' Leta upp kopia_excel_chart
    For Each shp In sld.Shapes
        If shp.Name = "kopia_excel_chart" Then
            Set chartShape = shp
            Exit For
        End If
    Next shp
    
    ' Leta upp labels_table
    For Each shp In sld.Shapes
        If shp.Name = "labels_table" Then
            Set tableShape = shp
            Exit For
        End If
    Next shp
    
 
    
    If Not tableShape Is Nothing Then
        tableShape.Delete
        Debug.Print "'labels_table' har tagits bort från sliden."
    Else
        Debug.Print "'labels_table' hittades inte."
    End If

End Sub

