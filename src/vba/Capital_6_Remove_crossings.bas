Attribute VB_Name = "Capital_6_Remove_crossings"
Sub Mac_Cap_Remove_Crossings()
    Dim pptSlide As slide
    Dim chartShape As shape

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Hitta diagrammet p� sliden
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades p� sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Ta bort crossings-linjer i diagrammet
    With chartShape.chart
        .Axes(xlCategory).CrossesAt = xlMinimum
        .Axes(xlValue).CrossesAt = xlMinimum
    End With

 End Sub

