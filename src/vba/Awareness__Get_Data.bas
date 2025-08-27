Attribute VB_Name = "Awareness__Get_Data"
Sub Awareness_Get_Data()
    Dim response As VbMsgBoxResult
    Dim pptSlide As slide
    Dim shp As shape
    Dim hasChart As Boolean

    ' Fr�ga anv�ndaren om de �r p� r�tt slide
    response = MsgBox("Are you on Awareness chart slide?", vbYesNo + vbQuestion, "Bekr�fta")
    If response = vbNo Then Exit Sub

    ' Kolla om det finns ett diagram p� sliden
    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    On Error GoTo 0

    hasChart = False
    For Each shp In pptSlide.Shapes
        If shp.hasChart Then
            hasChart = True
            Exit For
        End If
    Next shp

    If Not hasChart Then
        MsgBox "N�got gick fel. Kontrollera att det finns ett diagram p� sliden.", vbExclamation, "Fel"
        Exit Sub
    End If

    ' K�r delmakron
    Application.Run "AA_Startup"
    DoEvents

    Application.Run "AA_Total_3series"
    DoEvents

    Application.Run "AW_9C"
    DoEvents
End Sub

