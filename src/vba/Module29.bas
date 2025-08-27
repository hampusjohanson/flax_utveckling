Attribute VB_Name = "Module29"
Option Explicit

Sub Scatter_fix_9()

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart
    Dim ser As Object
    Dim pt As Object
    Dim i As Long
    
    Set sld = ActiveWindow.View.slide
    
    ' Hitta vänstra diagrammet (det som inte är kopia_excel_chart)
    For Each shp In sld.Shapes
        If shp.Type = msoChart And shp.Name <> "kopia_excel_chart" Then
            Set cht = shp.chart
            Debug.Print "Hittade diagrammet som ska justeras: " & shp.Name
            Exit For
        End If
    Next shp
    
    If cht Is Nothing Then
        Debug.Print "Hittade inget mål-diagram."
        Exit Sub
    End If
    
    ' Gå igenom alla punkter och vänsterställ textramarna
    For Each ser In cht.SeriesCollection
        For i = 1 To ser.Points.count
            Set pt = ser.Points(i)
            If pt.HasDataLabel Then
                With pt.dataLabel.TextFrame2
                    .HorizontalAnchor = msoAnchorNone ' Ingen speciell ankarpunkt
                    .textRange.ParagraphFormat.Alignment = msoAlignLeft
                End With
                Debug.Print "Punkt " & i & " vänsterställd."
            Else
                Debug.Print "Punkt " & i & " saknar etikett."
            End If
        Next i
    Next ser
    
    Debug.Print "Alla etiketter vänsterställda."

End Sub

