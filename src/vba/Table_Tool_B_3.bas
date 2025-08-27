Attribute VB_Name = "Table_Tool_B_3"
Sub Table_Tool_B_3()
    Dim selectedShape As shape
    Dim similarity_text As String
    Dim similarity_score_text As String
    Dim similarity_label As String
    Dim percentage As Integer

    ' Ensure that something is selected and check for text box selection
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a shape.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the selected shape has a text frame and is a text box
    Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
    
    If Not selectedShape.HasTextFrame Then
        MsgBox "Please select a text box.", vbExclamation
        Exit Sub
    End If
    
   ' Ensure the text frame has text
If Not selectedShape.TextFrame.HasText Then
    selectedShape.TextFrame.textRange.text = " " ' Insert a space to initialize
End If

    
    ' Format similarity_score as percentage
    If IsNumeric(similarity_score) Then
        percentage = Round(similarity_score * 100) ' Convert to percentage and round
    Else
        MsgBox "Invalid similarity_score value", vbExclamation
        Exit Sub
    End If

    ' Construct the similarity score text
    similarity_score_text = percentage & "% "
    similarity_label = "similarity"
    
    ' Combine both parts of the text
    similarity_text = similarity_score_text & similarity_label
    
    ' Insert the text into the selected text box
    selectedShape.TextFrame.textRange.text = similarity_text

    ' Set font size for similarity_score (percentage) part
    With selectedShape.TextFrame.textRange
        ' Set the font size for "similarity score in %"
        .Characters(1, Len(similarity_score_text)).Font.size = 11
        
        ' Set the font size for "similarity" part
        .Characters(Len(similarity_score_text) + 1, Len(similarity_label)).Font.size = 8
    End With

    Debug.Print "Text inserted: " & similarity_text
End Sub

