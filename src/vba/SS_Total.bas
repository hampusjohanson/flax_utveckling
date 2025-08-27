Attribute VB_Name = "SS_Total"
Sub SS_Total()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you on a slide with SUBSTANSMATRIS?", vbYesNo + vbQuestion, "Check Slide")
    If response = vbNo Then
        MsgBox "Macro cancelled.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next ' Avoid breaking if a macro fails

    Application.Run "SS_1"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "Chart_Remove_series"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "SS_3"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "SS_2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "SS_4"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "DataLabels2"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
        Application.Run "ListMissingVisiblePoints_Final"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
           Application.Run "CloseChartExcel_MacSafe"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents
    
        Application.Run "WaitOneSecond"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents

    Application.Run "CloseChartExcel_MacSafe"
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
    DoEvents


    
End Sub

Sub CloseChartExcel_MacSafe()
    Dim pptSlide As slide
    Dim shp As shape
    Dim chartObj As Object
    Dim xlApp As Object
    Dim wb As Object
    Dim i As Long

    ' Gå till aktuell slide
    Set pptSlide = ActiveWindow.View.slide

    ' Leta upp ett diagram
    For Each shp In pptSlide.Shapes
        If shp.hasChart Then
            Set chartObj = shp.chart

            ' Aktivera och få tag i arbetsboken och Excel-applikationen
            chartObj.chartData.Activate
            Set wb = chartObj.chartData.Workbook
            Set xlApp = wb.Application

            ' Stäng arbetsboken utan att spara
            wb.Saved = True
            wb.Close SaveChanges:=False

            ' Om Excel inte har fler arbetsböcker – stäng Excel
            If xlApp.Workbooks.count = 0 Then
                xlApp.Quit
            End If

            Exit Sub
        End If
    Next shp
End Sub

Sub WaitOneSecond()
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop
End Sub

