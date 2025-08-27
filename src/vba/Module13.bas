Attribute VB_Name = "Module13"
Sub Clean_All_Col2()
    ' Variabler
    Dim pptSlide As slide
    Dim tableShape As shape
    Dim sourceTable As table
    Dim rowIndex As Integer

    ' H�mta aktiv slide
    Set pptSlide = ActiveWindow.View.slide

    ' Loopa igenom alla shapes p� sliden
    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            Set sourceTable = tableShape.table
            ' Rensa kolumn 2 fr�n rad 2 och ned�t
            For rowIndex = 2 To sourceTable.Rows.count
                sourceTable.cell(rowIndex, 2).shape.TextFrame.textRange.text = ""
            Next rowIndex
        End If
    Next tableShape

End Sub
