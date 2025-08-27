Attribute VB_Name = "fconditional"
Option Explicit

'--- Hardcoded Office constants to avoid early binding ---------------------
Private Const ppSelectionShapes As Long = 2
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0

'===  PUBLIC ENTRY POINT  =================================================
Public Sub ApplyConditionalFormatting()
    Dim shp As Object, tbl As Object
    Dim r As Long, c As Long
    Dim cellVal As String, t As String
    Dim numVal1 As Double, numVal2 As Double, cellParsed As Double
    Dim isNumericLike As Boolean, matchCondition As Boolean
    Dim op As String, searchText As String
    Dim chosenFill As Long, chosenFont As Long, chosenBold As Boolean
    Dim i As Long, ch As String

    '--- 1) Check if a table is selected -----------------------------------
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a table first.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not CallByName(shp, "HasTable", VbGet) Then
        MsgBox "The selected shape is not a table.", vbExclamation
        Exit Sub
    End If
    Set tbl = CallByName(shp, "Table", VbGet)

    '--- 2) Show the UserForm ----------------------------------------------
    With frmtestcolor
        .StartUpPosition = 0
        .left = 800
        .Top = 300
        .Show

        If .Tag = "CANCEL" Then Exit Sub

        chosenFill = .GetSelectedColor
        chosenFont = .GetSelectedFontColor
        chosenBold = .GetSelectedBold
        op = .cmbOperator.value

        If chosenFill = -2 And Not chosenBold Then Exit Sub

        Select Case op
            Case "Greater than", "Less than"
                numVal1 = ParseNumberStrict(.txtValue1.value)
            Case "Between"
                numVal1 = ParseNumberStrict(.txtValue1.value)
                numVal2 = ParseNumberStrict(.txtValue2.value)
            Case "Contains"
                searchText = Trim$(.txtValue1.value)
                If Len(searchText) = 0 Then Exit Sub
            Case Else
                Exit Sub
        End Select
    End With

    '--- 3) Loop through all table cells -----------------------------------
    For r = 1 To tbl.Rows.count
        For c = 1 To tbl.Columns.count
            cellVal = Trim$(tbl.cell(r, c).shape.TextFrame.textRange.text)
            If Len(cellVal) > 0 Then
                matchCondition = False

                Select Case op
                    Case "Greater than", "Less than", "Between"
                        t = Replace(cellVal, " ", "")
                        t = Replace(t, Chr(160), "")
                        isNumericLike = True
                        For i = 1 To Len(t)
                            ch = Mid(t, i, 1)
                            If Not (ch Like "[0-9]" Or ch = "%" Or ch = "." Or ch = ",") Then
                                isNumericLike = False
                                Exit For
                            End If
                        Next i

                        If isNumericLike Then
                            cellParsed = ParseNumberStrict(cellVal)
                            Select Case op
                                Case "Greater than": matchCondition = (cellParsed > numVal1)
                                Case "Less than": matchCondition = (cellParsed < numVal1)
                                Case "Between": matchCondition = (cellParsed >= numVal1 And cellParsed <= numVal2)
                            End Select
                        End If

                    Case "Contains"
                        matchCondition = (InStr(1, cellVal, searchText, vbTextCompare) > 0)
                End Select

                If matchCondition Then
                    With tbl.cell(r, c).shape
                        If chosenFill <> -2 Then
                            If chosenFill = -1 Then
                                .Fill.visible = msoFalse
                            Else
                                .Fill.visible = msoTrue
                                .Fill.ForeColor.RGB = chosenFill
                            End If
                        End If
                        If chosenFont <> -2 Then
                            .TextFrame.textRange.Font.color.RGB = chosenFont
                        End If
                        If chosenBold Then
                            .TextFrame.textRange.Font.Bold = msoTrue
                        End If
                    End With
                End If
            End If
        Next c
    Next r
End Sub

'===  HELPER FUNCTION: strict number/% parser  ============================
Public Function ParseNumberStrict(ByVal s As String) As Double
    Dim t As String
    On Error Resume Next

    t = Replace$(s, " ", "")
    t = Replace$(t, Chr$(160), "")
    t = Trim$(t)

    If right(t, 1) = "%" Then
        t = left(t, Len(t) - 1)
        t = Replace(t, ",", ".")
        If IsNumeric(t) Then ParseNumberStrict = CDbl(t) / 100
    Else
        t = Replace(t, ",", ".")
        If IsNumeric(t) Then ParseNumberStrict = CDbl(t)
    End If
End Function

