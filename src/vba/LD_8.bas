Attribute VB_Name = "LD_8"
Sub Mac_LD_7()
    Dim pptSlide As slide
    Dim shape As shape

    ' Hämta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Gå igenom alla objekt på sliden
    For Each shape In pptSlide.Shapes
        ' Kolla om objektet är en textruta och har namnet 'topp_text_ruta'
        If shape.Name = "topp_text_ruta" Then
            shape.Delete ' Ta bort textrutan
            Exit Sub ' Avsluta makrot när textrutan är raderad
        End If
    Next shape

    ' Om textrutan inte finns, visa ett meddelande
    MsgBox "Textrutan hittades inte.", vbExclamation
End Sub

