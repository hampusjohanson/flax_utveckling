Attribute VB_Name = "AA_Check_Rightie"
Function AA_Check_Rightie_Leftie() As Boolean
    Dim pptSlide As slide
    Dim shape As shape
    Dim foundRightie As Boolean
    Dim foundLeftie As Boolean

    ' Set active slide
    Set pptSlide = ActiveWindow.View.slide

    ' Initialize variables
    foundRightie = False
    foundLeftie = False

    ' Loop through all shapes on the slide
    For Each shape In pptSlide.Shapes
        If shape.Name = "Rightie" Then foundRightie = True
        If shape.Name = "Leftie" Then foundLeftie = True
    Next shape

    ' If either "Rightie" or "Leftie" is missing, stop execution
    If Not (foundRightie And foundLeftie) Then
        Debug.Print "? Missing 'Rightie' or 'Leftie'. Stopping macro chain."
        AA_Check_Rightie_Leftie = False
        Exit Function
    End If

    ' If both exist, return True
    Debug.Print "? 'Rightie' and 'Leftie' found. Proceeding with the macro chain."
    AA_Check_Rightie_Leftie = True
End Function

