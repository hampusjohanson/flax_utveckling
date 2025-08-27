Attribute VB_Name = "Abbrev_2"
Sub Abbrev_2()
    Dim pptSlide As slide
    Dim leftTable As table
    Dim rightTable As table
    Dim rowIndex As Integer
    Dim smallestRowHeight As Single
    Dim currentRowHeight As Single
    Dim inconsistentRowsLeft As String
    Dim inconsistentRowsRight As String
    Dim rowChanges As String

    ' Initialize variables
    inconsistentRowsLeft = ""
    inconsistentRowsRight = ""
    rowChanges = "" ' Initialize string to store row changes for both tables

    ' Get the current slide
    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    If pptSlide Is Nothing Then
        MsgBox "No active slide found. Please select a slide and try again.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Retrieve left and right tables
    Dim tableCount As Integer
    tableCount = 0

    Dim shape As shape
    For Each shape In pptSlide.Shapes
        If shape.HasTable Then
            tableCount = tableCount + 1
            If tableCount = 1 Then
                Set leftTable = shape.table
            ElseIf tableCount = 2 Then
                Set rightTable = shape.table
                Exit For
            End If
        End If
    Next shape

    If leftTable Is Nothing Then
        MsgBox "No left table found on the slide.", vbExclamation
        Exit Sub
    End If

    If rightTable Is Nothing Then
        MsgBox "No right table found on the slide.", vbExclamation
        Exit Sub
    End If

    ' Find the smallest row height in the left table
    If leftTable.Rows.count > 1 Then
        smallestRowHeight = leftTable.Rows(2).height
        For rowIndex = 2 To leftTable.Rows.count
            currentRowHeight = leftTable.Rows(rowIndex).height
            If currentRowHeight < smallestRowHeight Then
                smallestRowHeight = currentRowHeight
            End If
        Next rowIndex

        ' Check for inconsistencies and adjust font size
        For rowIndex = 2 To leftTable.Rows.count
            currentRowHeight = leftTable.Rows(rowIndex).height
            If Abs(currentRowHeight - smallestRowHeight) > 0.01 Then ' Allow small differences
                inconsistentRowsLeft = inconsistentRowsLeft & "Row " & rowIndex & " height: " & currentRowHeight & " (expected: " & smallestRowHeight & "). " & vbCrLf

                ' Reduce font size in column 2
                On Error Resume Next
                With leftTable.cell(rowIndex, 2).shape.TextFrame.textRange.Font
                    If .size > 5 Then
                        .size = .size - 0.5
                    End If
                End With
                On Error GoTo 0

                rowChanges = rowChanges & "Left table row " & rowIndex & " made smaller." & vbCrLf
            End If
        Next rowIndex
    End If

    ' Find the smallest row height in the right table
    If rightTable.Rows.count > 1 Then
        smallestRowHeight = rightTable.Rows(2).height
        For rowIndex = 2 To rightTable.Rows.count
            currentRowHeight = rightTable.Rows(rowIndex).height
            If currentRowHeight < smallestRowHeight Then
                smallestRowHeight = currentRowHeight
            End If
        Next rowIndex

        ' Check for inconsistencies and adjust font size
        For rowIndex = 2 To rightTable.Rows.count
            currentRowHeight = rightTable.Rows(rowIndex).height
            If Abs(currentRowHeight - smallestRowHeight) > 0.01 Then ' Allow small differences
                inconsistentRowsRight = inconsistentRowsRight & "Row " & rowIndex & " height: " & currentRowHeight & " (expected: " & smallestRowHeight & "). " & vbCrLf

                ' Reduce font size in column 2
                On Error Resume Next
                With rightTable.cell(rowIndex, 2).shape.TextFrame.textRange.Font
                    If .size > 5 Then
                        .size = .size - 0.5
                    End If
                End With
                On Error GoTo 0

                rowChanges = rowChanges & "Right table row " & rowIndex & " made smaller." & vbCrLf
            End If
        Next rowIndex
    End If

    ' Display results for both tables in one message box
    If rowChanges <> "" Then
        MsgBox "The following rows have inconsistent heights and were made smaller:" & vbCrLf & rowChanges, vbExclamation
    Else
        MsgBox "All rows from row 2 downward in both tables have consistent heights.", vbInformation
    End If
End Sub

