Attribute VB_Name = "labeling_wide"
Option Explicit

Public LabelNames() As String
Public LabelWidths() As Double

Sub GetWide()
    Dim sld As slide
    Dim shp As shape
    Dim chrt As chart
    Dim srs As series
    Dim lbl As dataLabel
    Dim i As Integer
    Dim lblText As String
    Dim lblWidth As Double
    Dim RealLabelCount As Integer ' Count of valid labels

    ' Get the active slide
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then Exit Sub

    ' Find the first chart on the slide
    For Each shp In sld.Shapes
        If shp.hasChart Then
            Set chrt = shp.chart
            If chrt.SeriesCollection.count > 0 Then
                Set srs = chrt.SeriesCollection(1) ' First series

                ' **Första genomgången: Räkna riktiga etiketter**
                RealLabelCount = 0
                For i = 1 To srs.Points.count
                    If srs.Points(i).HasDataLabel Then
                        lblText = Trim(LCase(srs.Points(i).dataLabel.text))
                        lblText = Replace(lblText, Chr(160), "") ' Ta bort icke-brytande mellanslag

                        ' Debug: Visa etiketten
                        Debug.Print "Checking Label " & i & ": [" & lblText & "]"

                        If lblText <> "false" And lblText <> "falskt" And Len(lblText) > 0 Then
                            RealLabelCount = RealLabelCount + 1
                        End If
                    End If
                Next i

                ' **Debug: Skriv ut totalen av riktiga etiketter**
                Debug.Print "Total Real Labels Found: " & RealLabelCount

                ' **Endast fortsätt om vi har riktiga etiketter**
                If RealLabelCount > 0 Then
                    ' Skapa arrayer för att lagra namn och bredder
                    ReDim LabelNames(1 To RealLabelCount)
                    ReDim LabelWidths(1 To RealLabelCount)

                    ' **Andra genomgången: Spara etiketter och bredder**
                    RealLabelCount = 0 ' Återställ räknaren för array-indexering
                    For i = 1 To srs.Points.count
                        If srs.Points(i).HasDataLabel Then
                            Set lbl = srs.Points(i).dataLabel
                            lblText = Trim(LCase(lbl.text))
                            lblText = Replace(lblText, Chr(160), "") ' Ta bort icke-brytande mellanslag
                            lblWidth = lbl.width

                            ' Endast spara riktiga etiketter
                            If lblText <> "false" And lblText <> "falskt" And Len(lblText) > 0 Then
                                RealLabelCount = RealLabelCount + 1
                                LabelNames(RealLabelCount) = lblText
                                LabelWidths(RealLabelCount) = lblWidth

                                ' Debug: Skriv ut etikett och bredd
                                Debug.Print "Label: " & lblText & " | Width: " & lblWidth
                            End If
                        End If
                    Next i
                Else
                    Debug.Print "? No Real Labels Found!"
                End If
            End If
        End If
    Next shp
End Sub

