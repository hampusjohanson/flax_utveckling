Attribute VB_Name = "F2"
Sub Slide_Numbering_1()
    Dim slide As slide
    Dim newSlide As slide
    Dim i As Integer
    Dim columnCount As Integer
    Dim columnWidth As Single
    Dim columnSpacing As Single
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim currentColumn As Integer
    Dim textBox As shape
    Dim currentText As String
    Dim itemsPerColumn As Integer
    Dim titleCount As Integer
    Dim shape As shape
    Dim addedShapes As Collection
    Dim titleShape As shape
    Dim tableShapes As Collection
    Dim includeHidden As VbMsgBoxResult
    Dim slideNumber As Integer

    ' Ask the user if hidden slides should be included
    includeHidden = MsgBox("Include hidden slides in the overview?", vbYesNoCancel + vbQuestion, "Slide Overview")

    ' If user selects Cancel, exit macro
    If includeHidden = vbCancel Then Exit Sub

    Set addedShapes = New Collection
    Set tableShapes = New Collection

    ' Create a new slide at the end with Blank layout
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.count + 1, ppLayoutBlank)

    ' Add title text
    Set titleShape = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 20, 500, 50)
    titleShape.TextFrame.textRange.text = "Slide Overview"
    titleShape.TextFrame.textRange.Font.size = 22
    titleShape.TextFrame.textRange.Font.Bold = msoTrue
    addedShapes.Add titleShape ' Store title in protected list

    ' Get slide dimensions
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight

    ' Configure columns
    columnCount = 3 ' Three columns
    columnSpacing = 15
    columnWidth = (slideWidth - (columnSpacing * (columnCount + 1))) / columnCount

    ' Count slides with titles (respecting hidden slide selection)
    titleCount = 0
    For Each slide In ActivePresentation.Slides
        If slide.Shapes.HasTitle Then
            If includeHidden = vbYes Or slide.SlideShowTransition.Hidden = msoFalse Then
                titleCount = titleCount + 1
            End If
        End If
    Next slide

    ' Calculate items per column
    itemsPerColumn = titleCount \ columnCount
    If titleCount Mod columnCount > 0 Then itemsPerColumn = itemsPerColumn + 1

    ' Distribute titles across columns using actual SlideIndex
    i = 1
    currentColumn = 0
    currentText = ""

    For Each slide In ActivePresentation.Slides
        If slide.Shapes.HasTitle Then
            ' Check if slide should be included
            If includeHidden = vbYes Or slide.SlideShowTransition.Hidden = msoFalse Then
                slideNumber = slide.SlideIndex ' Use actual slide number
                currentText = currentText & slideNumber & ". " & slide.Shapes.title.TextFrame.textRange.text & vbCrLf

                ' Move to next column if max reached
                If i Mod itemsPerColumn = 0 Then
                    Set textBox = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        columnSpacing + (currentColumn * (columnWidth + columnSpacing)), 90, _
                        columnWidth, slideHeight - 130)

                    With textBox.TextFrame.textRange
                        .text = currentText
                        .Font.size = 9
                        .ParagraphFormat.Bullet.visible = msoFalse
                        .ParagraphFormat.SpaceBefore = 0
                        .ParagraphFormat.SpaceAfter = 1
                    End With

                    addedShapes.Add textBox
                    currentText = ""
                    currentColumn = currentColumn + 1
                End If

                i = i + 1
            End If
        End If
    Next slide

    ' Add remaining text in the last column
    If currentText <> "" Then
        Set textBox = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            columnSpacing + (currentColumn * (columnWidth + columnSpacing)), 90, _
            columnWidth, slideHeight - 130)

        With textBox.TextFrame.textRange
            .text = currentText
            .Font.size = 9
            .ParagraphFormat.Bullet.visible = msoFalse
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 1
        End With

        addedShapes.Add textBox
    End If

    ' Identify and store references to tables
    For Each shape In newSlide.Shapes
        If shape.HasTable Then
            tableShapes.Add shape
            If tableShapes.count = 2 Then Exit For
        End If
    Next shape

    ' Remove all other objects except title and tables
    For Each shape In newSlide.Shapes
        If Not ShapeInCollection(shape, addedShapes) And Not ShapeInCollection(shape, tableShapes) Then
            shape.Delete
        End If
    Next shape
End Sub

' Function to check if a shape is in the collection
Function ShapeInCollection(targetShape As shape, shapeCollection As Collection) As Boolean
    Dim TempShape As shape
    ShapeInCollection = False
    For Each TempShape In shapeCollection
        If TempShape.id = targetShape.id Then
            ShapeInCollection = True
            Exit Function
        End If
    Next TempShape
End Function


