Attribute VB_Name = "LD_8"
Sub Mac_LD_7()
    Dim pptSlide As slide
    Dim shape As shape

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' G� igenom alla objekt p� sliden
    For Each shape In pptSlide.Shapes
        ' Kolla om objektet �r en textruta och har namnet 'topp_text_ruta'
        If shape.Name = "topp_text_ruta" Then
            shape.Delete ' Ta bort textrutan
            Exit Sub ' Avsluta makrot n�r textrutan �r raderad
        End If
    Next shape

    ' Om textrutan inte finns, visa ett meddelande
    MsgBox "Textrutan hittades inte.", vbExclamation
End Sub

