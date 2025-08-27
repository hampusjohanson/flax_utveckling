Attribute VB_Name = "Corr_3_Text"
Sub Corr_3_Text()
    Dim pptSlide As slide
    Dim shapeItem As shape
    Dim tableCount As Integer
    Dim textBox As shape
    Dim foundTextBox As Boolean
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    tableCount = 0
    
    ' Count tables on the slide
    For Each shapeItem In pptSlide.Shapes
        If shapeItem.HasTable Then
            tableCount = tableCount + 1
        End If
    Next shapeItem

    Debug.Print "Number of tables found: " & tableCount
    
    ' Find "Rubrik 2" and update text
    foundTextBox = False
    For Each textBox In pptSlide.Shapes
        If textBox.HasTextFrame Then
            If textBox.Name = "Rubrik 2" Then
                textBox.TextFrame.textRange.text = "Correlations to the " & tableCount & " strongest drivers of volume premium"
                Debug.Print "Updated 'Rubrik 2' with: " & tableCount
                foundTextBox = True
                Exit For
            End If
        End If
    Next textBox

    ' If textbox not found, log a message
    If Not foundTextBox Then
        Debug.Print "Rubrik 2 not found."
    End If
End Sub


