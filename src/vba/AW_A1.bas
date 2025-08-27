Attribute VB_Name = "AW_A1"
Sub AW_A1()
    On Error GoTo ExitCleanly

    Dim pptSlide As slide
    Dim tableShape As shape
    Dim shp As shape
    Dim hasChart As Boolean

    ' Get the current slide
    Set pptSlide = ActiveWindow.View.slide

    ' Kontrollera att det finns minst ett diagram på sliden
    hasChart = False
    For Each shp In pptSlide.Shapes
        If shp.hasChart Then
            hasChart = True
            Exit For
        End If
    Next shp

    If Not hasChart Then
        MsgBox "Make sure there is a chart on the slide.", vbExclamation, "Fel"
        GoTo ExitCleanly
    End If

    ' Skapa tabellen
    Set tableShape = pptSlide.Shapes.AddTable(22, 7) ' 22 rows, 7 columns
    tableShape.Name = "SOURCE"
    tableShape.left = 50
    tableShape.Top = 50

ExitCleanly:
    Exit Sub
End Sub

