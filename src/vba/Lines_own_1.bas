Attribute VB_Name = "Lines_own_1"
Sub Lines_3_Select_Single_Series_Once_MacSafe()

    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim series As series
    Dim seriesList As Collection
    Dim inputMsg As String
    Dim userInput As String
    Dim selectionIndex As Integer
    Dim lineWidthInput As String
    Dim lineWidth As Single
    Dim i As Integer
    Dim chartFound As Boolean
    Dim seriesCount As Integer
    
    Set pptSlide = ActiveWindow.View.slide
    Set seriesList = New Collection
    chartFound = False
    
    ' Hämta serier från första diagrammet på sliden
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            seriesCount = chartObject.SeriesCollection.count
            chartFound = True
            
            ' Ignorera sista serien
            If seriesCount > 1 Then seriesCount = seriesCount - 1
            
            For i = 1 To seriesCount
                seriesList.Add i
                inputMsg = inputMsg & seriesList.count & ": " & chartObject.SeriesCollection(i).Name & vbCrLf
            Next i
            
            Exit For ' Endast första diagrammet behövs
        End If
    Next chartShape
    
    If Not chartFound Then Exit Sub
    
    ' Visa Selecta-menyn
    inputMsg = "Select which series to modify:" & vbCrLf & vbCrLf & inputMsg & vbCrLf & "Enter number:"
    userInput = InputBox(inputMsg, "Select Series")
    
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then Exit Sub
    
    selectionIndex = CInt(userInput)
    
    If selectionIndex < 1 Or selectionIndex > seriesList.count Then Exit Sub
    
    ' Fråga om linjebredd
    lineWidthInput = InputBox("Enter desired line width for selected series (" & chartObject.SeriesCollection(seriesList(selectionIndex)).Name & "):", _
                               "Line Width", 1.75)
    
    If lineWidthInput = "" Then Exit Sub
    If Not IsNumeric(lineWidthInput) Then Exit Sub
    
    lineWidth = CSng(lineWidthInput)
    If lineWidth <= 0 Then Exit Sub
    
    ' Applicera linjebredd på alla diagram på sliden för den valda serien
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            If chartObject.SeriesCollection.count >= seriesList(selectionIndex) Then
                Set series = chartObject.SeriesCollection(seriesList(selectionIndex))
                On Error Resume Next
                With series.Format.line
                    .Weight = lineWidth
                End With
                On Error GoTo 0
            End If
        End If
    Next chartShape
    
    ' Kör efterföljande makron
    On Error Resume Next
    Application.Run "Lines_7"
    Application.Run "Lines_Set_Markers"
    On Error GoTo 0

End Sub


