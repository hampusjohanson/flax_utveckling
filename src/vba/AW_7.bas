Attribute VB_Name = "AW_7"
Sub ShrinkChartFromTop()
    Const reduction As Single = 1 ' Hur mycket höjden ska minska (i punkter)

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart

    Set sld = ActiveWindow.View.slide

    ' Leta upp första diagrammet
    For Each shp In sld.Shapes
        If shp.hasChart Then
            With shp
                If .height > reduction Then
                    .Top = .Top + reduction
                    .height = .height - reduction
                Else
                    MsgBox "Diagrammet är redan för litet för att minska mer.", vbInformation
                End If
            End With
            Exit Sub
        End If
    Next shp

    MsgBox "Inget diagram hittades på sliden.", vbExclamation
End Sub

Sub GrowChartFromTop()
    Const increase As Single = 1 ' Hur mycket höjden ska öka (i punkter)

    Dim sld As slide
    Dim shp As shape

    Set sld = ActiveWindow.View.slide

    ' Leta upp första diagrammet
    For Each shp In sld.Shapes
        If shp.hasChart Then
            With shp
                .Top = .Top - increase
                .height = .height + increase
            End With
            Exit Sub
        End If
    Next shp

    MsgBox "Inget diagram hittades på sliden.", vbExclamation
End Sub

