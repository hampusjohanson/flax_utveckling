Attribute VB_Name = "Corr_Bold_99"
Sub Corr_Bold_99()
    Dim pptSlide As slide
    Dim s As shape
    Dim rowIndex As Integer
    Dim tbl As table
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loop through all shapes to find all tables
    For Each s In pptSlide.Shapes
        If s.HasTable Then
            Set tbl = s.table
            
            ' Remove bold formatting in columns 1 and 2 for all rows
            If Not tbl Is Nothing Then
                For rowIndex = 1 To tbl.Rows.count
                    On Error Resume Next
                    tbl.cell(rowIndex, 1).shape.TextFrame.textRange.Font.Bold = msoFalse
                    tbl.cell(rowIndex, 2).shape.TextFrame.textRange.Font.Bold = msoFalse
                    On Error GoTo 0
                Next rowIndex
            End If
            
            Debug.Print "Removed bold formatting from: " & s.Name
        End If
    Next s

    Debug.Print "Bold formatting cleared from all tables."
End Sub

