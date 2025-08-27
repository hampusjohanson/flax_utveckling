Attribute VB_Name = "Module6"
Sub Mac_Cap_Remove_Crossings_And_Hide_Axes()
    Dim pptSlide As slide
    Dim chartShape As shape

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Hitta diagrammet på sliden
    On Error Resume Next
    For Each shape In pptSlide.Shapes
        If shape.hasChart Then
            Set chartShape = shape
            Exit For
        End If
    Next shape
    If chartShape Is Nothing Then
        MsgBox "Inget diagram hittades på sliden.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Ta bort crossings-linjer i diagrammet och dölj axlarna
    With chartShape.chart
        ' Remove crossings
        .Axes(xlCategory).CrossesAt = xlMinimum
        .Axes(xlValue).CrossesAt = xlMinimum
        
        ' Hide both the category and value axes
        If .HasAxis(xlCategory, xlPrimary) Then
            .Axes(xlCategory, xlPrimary).visible = False
        End If
        If .HasAxis(xlValue, xlPrimary) Then
            .Axes(xlValue, xlPrimary).visible = False
        End If
    End With

    MsgBox "Crossings removed and axes hidden.", vbInformation
End Sub

