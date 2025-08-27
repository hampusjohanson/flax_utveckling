Attribute VB_Name = "AW_7"
Sub ShrinkChartFromTop()
    Const reduction As Single = 1 ' Hur mycket h�jden ska minska (i punkter)

    Dim sld As slide
    Dim shp As shape
    Dim cht As chart

    Set sld = ActiveWindow.View.slide

    ' Leta upp f�rsta diagrammet
    For Each shp In sld.Shapes
        If shp.hasChart Then
            With shp
                If .height > reduction Then
                    .Top = .Top + reduction
                    .height = .height - reduction
                Else
                    MsgBox "Diagrammet �r redan f�r litet f�r att minska mer.", vbInformation
                End If
            End With
            Exit Sub
        End If
    Next shp

    MsgBox "Inget diagram hittades p� sliden.", vbExclamation
End Sub

Sub GrowChartFromTop()
    Const increase As Single = 1 ' Hur mycket h�jden ska �ka (i punkter)

    Dim sld As slide
    Dim shp As shape

    Set sld = ActiveWindow.View.slide

    ' Leta upp f�rsta diagrammet
    For Each shp In sld.Shapes
        If shp.hasChart Then
            With shp
                .Top = .Top - increase
                .height = .height + increase
            End With
            Exit Sub
        End If
    Next shp

    MsgBox "Inget diagram hittades p� sliden.", vbExclamation
End Sub

