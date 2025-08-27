Attribute VB_Name = "Scatter_fix_3"
' === Modul: Scatter_fix_3.3 ===

Option Explicit

Sub Scatter_fix_3()

    Dim pres As Presentation
    Dim sld As slide
    Dim shp As shape
    Dim foundSlide As slide
    Dim chartShape As shape
    Dim chartFound As Boolean
    
    Set pres = ActivePresentation
    chartFound = False
    
    ' Leta upp sliden som har texten "Scatter calculation"
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If InStr(1, shp.TextFrame.textRange.text, "Scatter calculation", vbTextCompare) > 0 Then
                        Set foundSlide = sld
                        Debug.Print "Hittade slide: " & sld.SlideIndex
                        Exit For
                    End If
                End If
            End If
        Next shp
        If Not foundSlide Is Nothing Then Exit For
    Next sld
    
    If foundSlide Is Nothing Then
        Debug.Print "Ingen slide med texten 'Scatter calculation' hittades."
        Exit Sub
    End If
    
    ' Leta upp själva chart-shapen
    For Each shp In foundSlide.Shapes
        If shp.Type = msoChart Or shp.Type = msoPlaceholder Then
            If shp.Type = msoChart Then
                Set chartShape = shp
                chartFound = True
                Debug.Print "Hittade msoChart: " & shp.Name
                Exit For
            ElseIf shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.ContainedType = msoChart Then
                    Set chartShape = shp
                    chartFound = True
                    Debug.Print "Hittade chart i Placeholder: " & shp.Name
                    Exit For
                End If
            End If
        End If
    Next shp
    
    If Not chartFound Then
        Debug.Print "Inget diagram hittades på sliden."
        Exit Sub
    End If
    
    ' Duplicera själva shape-objektet (inte bara chart-data)
    Dim currentSlide As slide
    Set currentSlide = ActiveWindow.View.slide
    
    chartShape.Copy
    Dim pastedShape As shape
    Set pastedShape = currentSlide.Shapes.Paste(1)
    
    pastedShape.left = 800
    pastedShape.Top = 50
    pastedShape.Name = "kopia_excel_chart"
    
    Debug.Print "Kopia från 'Scatter calculation' skapad."

End Sub

