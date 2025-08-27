Attribute VB_Name = "SUB_TABLE_SCENARIO_1b"
Sub SUB_TABLE_SCENARIO_1b()
    ' Variables
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim targetTable As table

    ' === Find TARGET table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            If tableShape.Name = "TARGET" Then
                Set targetTable = tableShape.table
                Exit For
            End If
        End If
    Next tableShape

    ' If TARGET table is not found
    If targetTable Is Nothing Then
        Debug.Print "Table 'TARGET' not found on the slide."
        Exit Sub
    End If

    ' === Change fill color of Column 5, Row 1 to #FEBE61 ===
    targetTable.cell(1, 5).shape.Fill.ForeColor.RGB = RGB(254, 190, 97) ' #FEBE61

    Debug.Print "? SUB_TABLE_SCENARIO_1b EXECUTED! Column 5, Row 1 is now #FEBE61."
End Sub

