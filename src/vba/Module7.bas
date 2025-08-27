Attribute VB_Name = "Module7"
Sub FindAndDebugTable()
    Dim pptSlide As slide
    Dim shapeObj As shape
    Dim tableFound As Boolean

    Set pptSlide = ActiveWindow.View.slide
    tableFound = False

    Debug.Print "Checking shapes on the slide..."

    For Each shapeObj In pptSlide.Shapes
        Debug.Print "Shape Name: " & shapeObj.Name & " | Has Table: " & shapeObj.HasTable

        If shapeObj.Name = "TARGET" Then
            If shapeObj.HasTable Then
                Debug.Print "? Found Table in TARGET!"
                tableFound = True
                Exit For
            Else
                Debug.Print "? 'TARGET' does NOT contain a table!"
            End If
        End If
    Next shapeObj

    If Not tableFound Then
        Debug.Print "? Table not found inside 'TARGET'."
    End If
End Sub

