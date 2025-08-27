Attribute VB_Name = "Scatter_fix_7"

Option Explicit

Sub Scatter_fix_8()

    Dim sld As slide
    Dim shp As shape
    Dim kopiaChart As chart
    Dim targetChart As chart
    Dim wb As Object ' Workbook
    Dim ws As Object ' Worksheet
    Dim ser As Object
    Dim pt As Object
    Dim i As Long
    
    Set sld = ActiveWindow.View.slide
    
    ' === Steg 1: Hitta kopia_excel_chart f�r att l�sa koordinater ===
    For Each shp In sld.Shapes
        If shp.Name = "kopia_excel_chart" And shp.Type = msoChart Then
            Set kopiaChart = shp.chart
            Debug.Print "Hittade kopia_excel_chart"
            Exit For
        End If
    Next shp
    
    If kopiaChart Is Nothing Then
        Debug.Print "Hittade inte kopia_excel_chart."
        Exit Sub
    End If
    
    ' === Steg 2: Hitta det v�nstra (f�rsta) diagrammet p� sliden ===
    For Each shp In sld.Shapes
        If shp.Type = msoChart And shp.Name <> "kopia_excel_chart" Then
            Set targetChart = shp.chart
            Debug.Print "Hittade target chart: " & shp.Name
            Exit For
        End If
    Next shp
    
    If targetChart Is Nothing Then
        Debug.Print "Hittade inget m�l-diagram att uppdatera."
        Exit Sub
    End If
    
    ' === Steg 3: �ppna Excel i kopia_excel_chart ===
    Set wb = kopiaChart.chartData.Workbook
    
    On Error Resume Next
    Set ws = wb.Worksheets("New")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Debug.Print "Blad 'New' saknas."
        Exit Sub
    Else
        Debug.Print "Blad 'New' hittat i kopia_excel_chart"
    End If
    
    ' === Steg 4: G� till target chart och b�rja uppdatera ===
    If targetChart.SeriesCollection.count = 0 Then
        Debug.Print "Target chart har inga serier."
        Exit Sub
    End If
    
    Set ser = targetChart.SeriesCollection(1)
    
    For i = 1 To ser.Points.count
        Set pt = ser.Points(i)
        
        If pt.HasDataLabel Then
            
            Dim labelLeft As Variant, labelTop As Variant
            Dim labelWidth As Variant, labelHeight As Variant
            
            labelLeft = ws.Cells(i + 1, 3).value   ' C
            labelTop = ws.Cells(i + 1, 5).value    ' E
            labelWidth = ws.Cells(i + 1, 7).value  ' G
            labelHeight = ws.Cells(i + 1, 8).value ' H
            
            If IsNumeric(labelLeft) And IsNumeric(labelTop) And _
               IsNumeric(labelWidth) And IsNumeric(labelHeight) Then
               
                pt.dataLabel.left = labelLeft
                pt.dataLabel.Top = labelTop
                pt.dataLabel.width = labelWidth
                pt.dataLabel.height = labelHeight
                
                Debug.Print "Punkt " & i & " uppdaterad: Left=" & labelLeft & ", Top=" & labelTop
            Else
                Debug.Print "Punkt " & i & " saknar giltig data, hoppade �ver."
            End If
            
        Else
            Debug.Print "Punkt " & i & " saknar DataLabel."
        End If
        
    Next i
    
    Debug.Print "Alla etiketter f�rdiguppdaterade."
    
End Sub


