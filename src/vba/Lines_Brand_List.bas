Attribute VB_Name = "Lines_Brand_List"
Sub Lines_Brand_List()

 Application.Run "Lines_Legend_Delete_Tablesa"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

    Application.Run "Lines_Legend_New_Total_10"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents
    
     Application.Run "CopyAndSplit_BrandList_1_to_2"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents
    
     Application.Run "DeleteLegendArrowElements"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents
    
    Application.Run "CreateLegendArrowPair_Under_BrandList1"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents
    
End Sub


Sub Lines_Brand_List_OLD()
    On Error Resume Next ' Avoid breaking if a macro fails


    Application.Run "Lines_12"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

    Application.Run "Lines_13"
    If Err.Number <> 0 Then MsgBox "Error in SP_1: " & Err.Description
    DoEvents

    Application.Run "Lines_14"
    If Err.Number <> 0 Then MsgBox "Error in SP_2: " & Err.Description
    DoEvents

    Application.Run "Lines_15"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
    
    Application.Run "Lines_16"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
    
  Application.Run "Lines_17"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
    
    Application.Run "Lines_18"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
 Application.Run "Lines_21a"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
 
    
      Application.Run "Lines_22"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents

      Application.Run "Lines_23"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents

      Application.Run "Lines_24"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents

      Application.Run "Lines_25"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents

    Application.Run "Lines_26"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
   
      Application.Run "Lines_27"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
     Application.Run "Lines_21"
    If Err.Number <> 0 Then MsgBox "Error in SP_3: " & Err.Description
    DoEvents
   
End Sub


Sub CreateLegendArrowPair_Under_BrandList1()
    Dim pptSlide As slide
    Dim shape1 As shape
    Dim lineLeft As shape, lineRight As shape
    Dim yPos As Single
    Dim lineLength As Single: lineLength = 0.3 * 28.35
    Dim lineYgap As Single: lineYgap = 8
    Dim leftX As Single, rightX As Single

    Set pptSlide = ActiveWindow.View.slide

    On Error Resume Next
    Set shape1 = pptSlide.Shapes("Brand_List_1")
    On Error GoTo 0

    If shape1 Is Nothing Then
        MsgBox "Brand_List_1 hittades inte.", vbExclamation
        Exit Sub
    End If

    ' Vertikal position: strax under Brand_List_1
    yPos = shape1.Top + shape1.height + lineYgap

    ' === Högerpekande linje (vänster placering) ===
    leftX = 28.4 * 26.35
    Set lineRight = pptSlide.Shapes.AddLine(BeginX:=leftX, BeginY:=yPos, endX:=leftX + lineLength, endY:=yPos)
    lineRight.Name = "line_own_1"

    With lineRight.line
        .ForeColor.RGB = RGB(17, 21, 66)
        .Weight = 1.75
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOval
        .EndArrowheadLength = msoArrowheadLengthMedium
        .EndArrowheadWidth = msoArrowheadWidthMedium
        .DashStyle = msoLineSolid
        .visible = msoTrue
    End With

    ' === Vänsterpekande linje (höger placering, ritas höger – sen roteras) ===
    rightX = 28.7 * 26.35
    Set lineLeft = pptSlide.Shapes.AddLine(BeginX:=rightX, BeginY:=yPos, endX:=rightX + lineLength, endY:=yPos)
    lineLeft.Name = "line_own_2"

    With lineLeft.line
        .ForeColor.RGB = RGB(17, 21, 66)
        .Weight = 1.75
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOval
        .EndArrowheadLength = msoArrowheadLengthMedium
        .EndArrowheadWidth = msoArrowheadWidthMedium
        .DashStyle = msoLineSolid
        .visible = msoTrue
    End With

    lineLeft.Rotation = 180

    ' === Avgör språk baserat på slide-text ===
    Dim shapeText As shape
    Dim textToAdd As String
    Dim foundText As String: foundText = ""

    For Each shapeText In pptSlide.Shapes
        If shapeText.HasTextFrame Then
            If shapeText.TextFrame.HasText Then
                foundText = foundText & shapeText.TextFrame.textRange.text
            End If
        End If
    Next shapeText

    If InStr(foundText, "Stronger") > 0 Then
        textToAdd = "Ownership"
    ElseIf InStr(foundText, "Starkare") > 0 Then
        textToAdd = "Ägarskap"
    Else
        textToAdd = ""
    End If

    ' === Skapa textbox om vi har något att skriva ===
    If textToAdd <> "" Then
        Dim textShape As shape
        Dim textLeft As Single
        textLeft = lineLeft.left + lineLeft.width + (0.01 * 28.35)

        Set textShape = pptSlide.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            left:=textLeft, _
            Top:=lineLeft.Top - 8.3, _
            width:=3 * 28.35, _
            height:=20)

        With textShape.TextFrame.textRange
            .text = textToAdd
            .Font.Name = "Arial"
            .Font.size = 8
            .Font.color.RGB = RGB(17, 21, 66)
        End With

        textShape.Name = "line_own_label"
    End If

    Debug.Print "Linjer och ev. etikett skapade."
End Sub

Sub DeleteLegendArrowElements()
    Dim pptSlide As slide
    Dim sName As Variant
    Dim shapeObj As shape

    Set pptSlide = ActiveWindow.View.slide

    For Each sName In Array("line_own_1", "line_own_2", "line_own_label")
        On Error Resume Next
        Set shapeObj = pptSlide.Shapes(sName)
        If Not shapeObj Is Nothing Then
            shapeObj.Delete
            Debug.Print sName & " deleted."
        End If
        Set shapeObj = Nothing
        On Error GoTo 0
    Next sName
End Sub

Sub CopyAndSplit_BrandList_1_to_2()
    Dim pptSlide As slide
    Dim chartShape As shape
    Dim chartObject As chart
    Dim sIndex As Integer
    Dim brand_count As Integer
    Dim tbl1 As shape, tbl2 As shape
    Dim rowCount As Integer
    Dim keepInTbl1 As Integer, keepInTbl2 As Integer
    Dim i As Integer

    Set pptSlide = ActiveWindow.View.slide

    ' === Räkna antal synliga serier i första diagrammet ===
    brand_count = 0
    For Each chartShape In pptSlide.Shapes
        If chartShape.hasChart Then
            Set chartObject = chartShape.chart
            For sIndex = 1 To chartObject.SeriesCollection.count
                If chartObject.SeriesCollection(sIndex).Format.line.visible = msoTrue Then
                    brand_count = brand_count + 1
                End If
            Next sIndex
            Exit For
        End If
    Next chartShape

    ' Hoppa över om färre än 5
    If brand_count < 5 Then
        On Error Resume Next
        pptSlide.Shapes("Brand_List_2").Delete
        On Error GoTo 0
        Debug.Print "Färre än 5 varumärken – Brand_List_2 skapas inte."
        Exit Sub
    End If

    ' === Hämta Brand_List_1 ===
    On Error Resume Next
    Set tbl1 = pptSlide.Shapes("Brand_List_1")
    If tbl1 Is Nothing Then
        MsgBox "Brand_List_1 finns inte!", vbExclamation
        Exit Sub
    End If

    ' Ta bort gammal Brand_List_2 om den finns
    pptSlide.Shapes("Brand_List_2").Delete
    On Error GoTo 0

    ' === Kopiera tabell ===
    tbl1.Copy
    pptSlide.Shapes.Paste.Name = "Brand_List_2"
    Set tbl2 = pptSlide.Shapes("Brand_List_2")

    ' Placera Brand_List_2 till höger om Brand_List_1
    tbl2.Top = tbl1.Top
    tbl2.left = 28.5 * 29.65

    ' === Dela upp rader ===
    rowCount = tbl1.table.Rows.count

    If rowCount Mod 2 = 0 Then
        keepInTbl1 = rowCount \ 2
        keepInTbl2 = rowCount \ 2
    Else
        keepInTbl1 = (rowCount \ 2) + 1
        keepInTbl2 = rowCount - keepInTbl1
    End If

    ' Ta bort nedersta rader från Brand_List_1
    For i = rowCount To (keepInTbl1 + 1) Step -1
        tbl1.table.Rows(i).Delete
    Next i

    ' Ta bort översta rader från Brand_List_2
    For i = 1 To keepInTbl1
        tbl2.table.Rows(1).Delete
    Next i

    Debug.Print "Brand_List_1 behåller " & keepInTbl1 & " rader, Brand_List_2 får " & keepInTbl2 & " rader."
End Sub

