Attribute VB_Name = "Capital_913_Remove_Balance"
Sub Mac_Cap_Balance_Remove()
    Dim pptSlide As slide
    Dim shape As shape
    Dim loopAgain As Boolean

    ' H�mta aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    Do
        loopAgain = False ' �terst�ll loopflagga

        ' S�k efter formerna med namnen "Balance" och "Balance_Text" och ta bort dem
        On Error Resume Next
        For Each shape In pptSlide.Shapes
            If shape.Name = "BalanceCircle" Or shape.Name = "MyTextBox" Then
                shape.Delete
                loopAgain = True ' Om n�got tas bort, k�r loopen igen
            End If
        Next shape
        On Error GoTo 0

    Loop While loopAgain ' K�r tills det inte finns fler f�rekomster

End Sub


