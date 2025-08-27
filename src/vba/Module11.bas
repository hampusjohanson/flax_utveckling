Attribute VB_Name = "Module11"
Sub ReportSlidesWithMissingOrDeletedNumbers()
    Dim slide As slide
    Dim shape As shape
    Dim slideNum As Integer
    Dim numberFound As Boolean
    Dim potentialNumberNames As Variant
    Dim i As Integer
    Dim layoutName As String
    Dim issueReport As String

    ' Lista över möjliga numreringsnamn på olika språk
    potentialNumberNames = Array("slide number", "bildnummer", "page", "num", "sida", "platshållare för")

    ' Loopa genom alla slides och hitta numrering
    For Each slide In ActivePresentation.Slides
        slideNum = slide.SlideIndex
        numberFound = False
        layoutName = LCase(slide.CustomLayout.Name)

        ' Hoppa över slides där layouten börjar med "Chapter", "Title", "Rubrikbild" eller "Start"
        If left(layoutName, 7) = "chapter" Or left(layoutName, 5) = "title" Or _
           left(layoutName, 10) = "rubrikbild" Or left(layoutName, 5) = "start" Then GoTo NextSlide

        ' **Kolla om sliden HAR en numreringsplatshållare**
        Dim hasPlaceholder As Boolean
        hasPlaceholder = False
        
        For Each shape In slide.Shapes
            If shape.Type = msoPlaceholder Then
                ' Kontrollera om det är en numreringsplatshållare (baserat på namn)
                For i = LBound(potentialNumberNames) To UBound(potentialNumberNames)
                    If InStr(1, LCase(shape.Name), potentialNumberNames(i), vbTextCompare) > 0 Then
                        hasPlaceholder = True
                        Exit For
                    End If
                Next i
            End If
        Next shape

        ' **Kolla om numrering faktiskt finns**
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    For i = LBound(potentialNumberNames) To UBound(potentialNumberNames)
                        If InStr(1, LCase(shape.Name), potentialNumberNames(i), vbTextCompare) > 0 Then
                            numberFound = True
                            Exit For
                        End If
                    Next i
                End If
            End If
            If numberFound Then Exit For
        Next shape

        ' **Om sliden skulle ha haft en numrering men den saknas, flagga den**
        If hasPlaceholder And Not numberFound Then
            issueReport = issueReport & "Slide " & slideNum & ": Numrering har tagits bort." & vbCrLf
        End If

NextSlide:
    Next slide

    ' ?? **Slutrapport**
    If issueReport = "" Then
        MsgBox "Alla slides som ska ha numrering har det!", vbInformation
    Else
        MsgBox "Följande slides saknar numrering som borde finnas:" & vbCrLf & issueReport, vbExclamation
    End If
End Sub


