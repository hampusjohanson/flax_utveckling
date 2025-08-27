Attribute VB_Name = "Chapter_Module"
' === Modul: Chapter_Module ===

Public gChapterBold As Boolean
Public gChapterItalic As Boolean
Public gChapterAskEach As Boolean
Public gChapterFontName As String
Public gChapterFontSize As Single
Public gChapterAbort As Boolean
Public gChapterSettingsConfirmed As Boolean

Sub ApplyChapterHeadersWithUserForm()
    Dim i As Integer
    Dim sld As slide
    Dim shp As shape
    Dim slideWidth As Single
    Dim candidateShape As shape
    Dim fullText As String
    Dim endSlideReached As Boolean

    PrepareChapters
    gChapterAbort = False
    gChapterSettingsConfirmed = False
    frmChapterSettings.Show
    If gChapterAbort Or Not gChapterSettingsConfirmed Then Exit Sub

    slideWidth = ActivePresentation.PageSetup.slideWidth
    endSlideReached = False

    For i = 1 To UBound(ChapterDataModule.ChapterList)
        Dim fromSld As Integer: fromSld = ChapterDataModule.ChapterList(i).SlideFrom
        Dim toSld As Integer: toSld = ChapterDataModule.ChapterList(i).SlideTo
        If fromSld = 0 Or toSld = 0 Then GoTo SkipRange

        If ChapterDataModule.ChapterList(i).HeadlineBold <> "" Or ChapterDataModule.ChapterList(i).HeadlineText <> "" Then
            fullText = ChapterDataModule.ChapterList(i).HeadlineBold & ChapterDataModule.ChapterList(i).HeadlineText
        Else
            fullText = ToSentenceCase(ChapterDataModule.ChapterList(i).DividerText)
        End If

        For Each sld In ActivePresentation.Slides
            If sld.SlideIndex >= fromSld And sld.SlideIndex <= toSld Then

                If sld.CustomLayout.Name = "Start-/End slide" Then
                    endSlideReached = True
                    Exit For
                End If

                If endSlideReached Then Exit For

                Set candidateShape = Nothing

                For Each shp In sld.Shapes
                    If shp.HasTextFrame Then
                        If shp.left < slideWidth / 2 And shp.Top < 100 And shp.height < 80 Then
                            If candidateShape Is Nothing Then
                                Set candidateShape = shp
                            ElseIf shp.Top < candidateShape.Top Then
                                Set candidateShape = shp
                            End If
                        End If
                    End If
                Next shp

                If Not candidateShape Is Nothing Then
                    If gChapterAskEach Then
                        sld.Select
                        ActiveWindow.View.GotoSlide sld.SlideIndex
                        DoEvents

                        Dim userResp As VbMsgBoxResult
                        userResp = MsgBox( _
                            "Slide " & sld.SlideIndex & vbNewLine & _
                            "Apply this chapter title?" & vbNewLine & _
                            """" & fullText & """", _
                            vbYesNoCancel + vbQuestion)

                        If userResp = vbCancel Then
                            gChapterAbort = True
                            Exit Sub
                        ElseIf userResp = vbNo Then
                            GoTo SkipThisSlide
                        End If
                    End If

                    With candidateShape.TextFrame.textRange
                        .text = fullText
                        If ChapterDataModule.ChapterList(i).HeadlineBold <> "" And ChapterDataModule.ChapterList(i).HeadlineText <> "" Then
                            .Font.Bold = False
                            .Font.Italic = False
                            .Font.Name = gChapterFontName
                            .Font.size = gChapterFontSize
                            .Characters(1, Len(ChapterDataModule.ChapterList(i).HeadlineBold)).Font.Bold = True
                        Else
                            .Font.Bold = gChapterBold
                            .Font.Italic = gChapterItalic
                            .Font.Name = gChapterFontName
                            .Font.size = gChapterFontSize
                        End If
                    End With
                End If
SkipThisSlide:
            End If
        Next sld
SkipRange:
    Next i

    If Not gChapterAbort Then
        MsgBox "Chapter titles applied.", vbInformation
    End If
End Sub

Function ToSentenceCase(ByVal txt As String) As String
    txt = Trim(txt)
    If Len(txt) = 0 Then
        ToSentenceCase = ""
    Else
        ToSentenceCase = UCase(left(txt, 1)) & LCase(Mid(txt, 2))
    End If
End Function


