Attribute VB_Name = "Check_Data_Points"
Sub ListMissingVisiblePoints_Final()
    Dim pptSlide As slide
    Dim shp As shape
    Dim chartObj As Object
    Dim srs As Object
    Dim pt As Object
    Dim i As Long
    Dim wb As Object
    Dim ws As Object
    Dim xlApp As Object
    Dim labelText As String
    Dim totalVisible As Long
    Dim ptLeft As Double, ptTop As Double
    Dim paLeft As Double, paTop As Double, paWidth As Double, paHeight As Double
    Dim inBounds As Boolean
    Dim validLabels As Collection
    Dim visibleLabels As Collection
    Dim labelKey As String
    Dim cellVal As String
    Dim missingOutput As String
    Dim itm As Variant

    Set pptSlide = ActiveWindow.View.slide

    For Each shp In pptSlide.Shapes
        If shp.hasChart Then
            Set chartObj = shp.chart
            chartObj.chartData.Activate

            Set wb = chartObj.chartData.Workbook
            Set ws = wb.Worksheets(1)
            Set xlApp = wb.Application

            ' Hämta plot area-gränser
            With chartObj.PlotArea
                paLeft = .InsideLeft
                paTop = .InsideTop
                paWidth = .InsideWidth
                paHeight = .InsideHeight
            End With

            ' Samla alla giltiga etiketter från kolumn A
            Set validLabels = New Collection
            i = 2
            Do While ws.Cells(i, 1).value <> ""
                cellVal = Trim(ws.Cells(i, 1).text)
                If LCase(cellVal) <> "false" And LCase(cellVal) <> "falskt" And cellVal <> "" Then
                    On Error Resume Next
                    validLabels.Add cellVal, cellVal
                    On Error GoTo 0
                End If
                i = i + 1
            Loop

            ' Samla synliga etiketter inom plot area
            Set visibleLabels = New Collection
            totalVisible = 0

            For Each srs In chartObj.SeriesCollection
                For i = 1 To srs.Points.count
                    Set pt = srs.Points(i)

                    On Error Resume Next
                    If pt.Format.Fill.visible Then
                        ptLeft = pt.left
                        ptTop = pt.Top

                        inBounds = (ptLeft >= paLeft And ptLeft <= paLeft + paWidth) And _
                                   (ptTop >= paTop And ptTop <= paTop + paHeight)

                        If inBounds Then
                            labelText = Trim(ws.Cells(i + 1, 1).text)
                            If labelText <> "" Then
                                On Error Resume Next
                                visibleLabels.Add labelText, labelText
                                On Error GoTo 0
                                totalVisible = totalVisible + 1
                            End If
                        End If
                    End If
                    On Error GoTo 0
                Next i
            Next srs

            ' Jämför och identifiera saknade
            missingOutput = ""
            For Each itm In validLabels
                On Error Resume Next
                labelKey = visibleLabels.Item(itm)
                If Err.Number <> 0 Then
                    missingOutput = missingOutput & "- " & itm & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
            Next itm

            ' Tvinga stäng Excel-instansen PowerPoint öppnade
            xlApp.DisplayAlerts = False
            xlApp.Quit

            ' Visa om något saknas
            If missingOutput <> "" Then
                MsgBox "OBS: " & validLabels.count - totalVisible & " saknas visuellt i graf:" & vbCrLf & missingOutput, _
                    vbExclamation, "Varning: Dolda etiketter"
            End If

            Exit Sub
        End If
    Next shp

    MsgBox "Inget diagram hittades på sliden.", vbExclamation
End Sub


