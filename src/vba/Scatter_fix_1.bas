Attribute VB_Name = "Scatter_fix_1"
' === Modul: Scatter_fix_1 ===

Option Explicit

Sub Scatter_fix_1()

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim ser As Object
    Dim pt As Object
    Dim i As Integer
    Dim chartCount As Integer
    Dim totalPoints As Integer
    Dim validLabels As Integer
    
    Dim labelLeft As Single
    Dim labelTop As Single
    Dim labelWidth As Single
    Dim labelHeight As Single
    Dim labelRight As Single
    Dim labelBottom As Single
    Dim labelText As String
    
    Dim tableShape As shape
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim rowIndex As Integer
    
    ' Tabell med 50 rader + header, 8 kolumner (utan Valid Label)
    rowCount = 51
    colCount = 8
    
    chartCount = 0
    validLabels = 0
    
    Set sld = ActiveWindow.View.slide
    
    Debug.Print "Slide: " & sld.SlideIndex
    
    ' Skapa tabellen på sliden
    Set tableShape = sld.Shapes.AddTable(rowCount, colCount, 50, 50, 700, 400)
    tableShape.Name = "labels_table"
    
    ' Sätt tabellrubriker
    With tableShape.table
        .cell(1, 1).shape.TextFrame.textRange.text = "#"
        .cell(1, 2).shape.TextFrame.textRange.text = "Text"
        .cell(1, 3).shape.TextFrame.textRange.text = "Left"
        .cell(1, 4).shape.TextFrame.textRange.text = "Right"
        .cell(1, 5).shape.TextFrame.textRange.text = "Top"
        .cell(1, 6).shape.TextFrame.textRange.text = "Bottom"
        .cell(1, 7).shape.TextFrame.textRange.text = "Width"
        .cell(1, 8).shape.TextFrame.textRange.text = "Height"
    End With
    
    ' Fyll radernas index
    For rowIndex = 2 To rowCount
        tableShape.table.cell(rowIndex, 1).shape.TextFrame.textRange.text = CStr(rowIndex - 1)
    Next rowIndex
    
    ' Börja loopa alla shapes
    For Each shp In sld.Shapes
        If shp.Type = msoChart Then
            chartCount = chartCount + 1
            Debug.Print "Found Chart #" & chartCount & ": " & shp.Name
            
            Set cht = shp.chart
            
            Select Case cht.chartType
                Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                    Debug.Print "  Chart type: Scatter"
                    
                    For Each ser In cht.SeriesCollection
                        Debug.Print "    Series: " & ser.Name
                        
                        totalPoints = ser.Points.count
                        Debug.Print "    Total points in series: " & totalPoints
                        
                        For i = 1 To totalPoints
                            Set pt = ser.Points(i)
                            
                            If pt.HasDataLabel Then
                                On Error Resume Next
                                If Not pt.dataLabel Is Nothing Then
                                    If IsValidNumber(pt.dataLabel.left) Then
                                        
                                        labelLeft = pt.dataLabel.left
                                        labelTop = pt.dataLabel.Top
                                        labelWidth = pt.dataLabel.width
                                        labelHeight = pt.dataLabel.height
                                        labelRight = labelLeft + labelWidth
                                        labelBottom = labelTop + labelHeight
                                        labelText = pt.dataLabel.text
                                        
                                        validLabels = validLabels + 1
                                        
                                        Debug.Print "Valid Label #" & validLabels & ": " & labelText
                                        
                                        ' Fyll tabellen
                                        tableShape.table.cell(i + 1, 2).shape.TextFrame.textRange.text = labelText
                                        tableShape.table.cell(i + 1, 3).shape.TextFrame.textRange.text = Format(labelLeft, "0.00")
                                        tableShape.table.cell(i + 1, 4).shape.TextFrame.textRange.text = Format(labelRight, "0.00")
                                        tableShape.table.cell(i + 1, 5).shape.TextFrame.textRange.text = Format(labelTop, "0.00")
                                        tableShape.table.cell(i + 1, 6).shape.TextFrame.textRange.text = Format(labelBottom, "0.00")
                                        tableShape.table.cell(i + 1, 7).shape.TextFrame.textRange.text = Format(labelWidth, "0.00")
                                        tableShape.table.cell(i + 1, 8).shape.TextFrame.textRange.text = Format(labelHeight, "0.00")
                                        
                                    Else
                                        Debug.Print "Point #" & i & " has invalid (NaN) coordinates."
                                    End If
                                Else
                                    Debug.Print "Point #" & i & " DataLabel object is Nothing."
                                End If
                                On Error GoTo 0
                            Else
                                Debug.Print "Point #" & i & " has no data label."
                            End If
                        Next i
                        
                        Debug.Print "Total valid labels found: " & validLabels
                        
                    Next ser
                Case Else
                    Debug.Print "Chart type: Not scatter (skipped)"
            End Select
            
        End If
    Next shp
    
    If chartCount = 0 Then
        Debug.Print "No charts found on this slide."
    End If
    
End Sub

' === Vattentät NaN-kontroll ===

Function IsValidNumber(value As Variant) As Boolean
    On Error Resume Next
    Dim txt As String
    txt = CStr(value)
    If InStr(1, txt, "nan", vbTextCompare) > 0 Then
        IsValidNumber = False
    Else
        IsValidNumber = IsNumeric(value)
    End If
    On Error GoTo 0
End Function

