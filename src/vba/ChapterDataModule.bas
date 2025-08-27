Attribute VB_Name = "ChapterDataModule"
' === Modul: ChapterDataModule ===

Public Type ChapterEntry
    SlideIndex As Integer
    DividerText As String
    HeadlineBold As String
    HeadlineText As String
    SlideFrom As Integer
    SlideTo As Integer
End Type

Public ChapterList() As ChapterEntry

Sub OpenChapterEditor()
    PrepareChapters
    Chapterbox2.Show
End Sub

Public Sub PrepareChapters()
    Dim sld As slide
    Dim shp As shape
    Dim i As Long
    Dim currentTitle As String
    Dim currentSlide As Integer
    Dim dividerIndices() As Integer
    Dim foundCount As Integer

    For Each sld In ActivePresentation.Slides
        If sld.CustomLayout.Name Like "Chapter*" Or sld.CustomLayout.Name Like "Title Slide*" Then
            foundCount = foundCount + 1
            ReDim Preserve dividerIndices(1 To foundCount)
            dividerIndices(foundCount) = sld.SlideIndex
        End If
    Next sld

    ReDim ChapterList(1 To foundCount)

    For i = 1 To foundCount
        currentSlide = dividerIndices(i)
        currentTitle = ""

        For Each shp In ActivePresentation.Slides(currentSlide).Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    If Len(shp.TextFrame.textRange.text) > 0 Then
                        If shp.TextFrame.textRange.Font.size >= 20 Then
                            currentTitle = Trim(shp.TextFrame.textRange.text)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next shp

        ChapterList(i).SlideIndex = currentSlide
        ChapterList(i).DividerText = currentTitle
        ChapterList(i).HeadlineText = ToSentenceCase(currentTitle)
        ChapterList(i).HeadlineBold = ""
        ChapterList(i).SlideFrom = currentSlide + 1

        If i < foundCount Then
            If dividerIndices(i + 1) = currentSlide + 1 Then
                ChapterList(i).SlideTo = 0
            Else
                ChapterList(i).SlideTo = dividerIndices(i + 1) - 1
            End If
        Else
            ChapterList(i).SlideTo = ActivePresentation.Slides.count
        End If
    Next i
End Sub

Public Function ToSentenceCase(ByVal txt As String) As String
    txt = Trim(txt)
    If Len(txt) = 0 Then
        ToSentenceCase = ""
    Else
        ToSentenceCase = UCase(left(txt, 1)) & LCase(Mid(txt, 2))
    End If
End Function

