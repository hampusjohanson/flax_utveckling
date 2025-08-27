Attribute VB_Name = "Background_1_7"
Sub Background_1_7()
    ' Step 1: Read Stronger_Last_Value and Associations_Total from the CSV
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim currentRow As Long
    Dim dataArray() As String
    Dim Stronger_Last_Value As Double
    Dim Associations_Total As Double

    ' Dynamically get the file path to the desktop
    If Environ("OS") Like "*Windows*" Then
        filePath = "c:\Local\exported_data_semi.csv"
    Else
        filePath = "/Users/" & Environ("USER") & "/Desktop/exported_data_semi.csv"
    End If

    ' Debug output
    Debug.Print "File path: " & filePath

    ' Check if file exists
    If Dir(filePath) = "" Then
        Debug.Print "File not found: " & filePath
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Open the CSV file and read Stronger_Last_Value and Associations_Total
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    currentRow = 0

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        currentRow = currentRow + 1
        
        If currentRow = 469 Then ' Associations row
            dataArray = Split(lineData, ";")
            Associations_Total = CDbl(dataArray(1))
            Debug.Print "Associations_Total read: " & Associations_Total
        End If
        
        If currentRow = 470 Then ' Stronger_Last_Value row
            dataArray = Split(lineData, ";")
            Stronger_Last_Value = CDbl(dataArray(1))
            Debug.Print "Stronger_Last_Value read: " & Stronger_Last_Value
            Exit Do
        End If
    Loop
    Close #fileNumber

    ' Step 2: Reference the slide and find the tables named LEFTIE and RIGHTIE
    Dim pptSlide As slide
    Set pptSlide = ActiveWindow.View.slide
    Debug.Print "Slide reference set."

    Dim pptTableLeft As table, pptTableRight As table
    Dim s As shape
    For Each s In pptSlide.Shapes
        If s.Name = "LEFTIE" And s.HasTable Then
            Set pptTableLeft = s.table
        ElseIf s.Name = "RIGHTIE" And s.HasTable Then
            Set pptTableRight = s.table
        End If
    Next s

    If pptTableLeft Is Nothing Or pptTableRight Is Nothing Then
        Debug.Print "One or both tables not found."
        MsgBox "Tables LEFTIE or RIGHTIE not found on slide.", vbExclamation
        Exit Sub
    End If

    ' Step 3: Apply color formatting to column 3 and 4 from row 3 downwards in both tables
    Dim i As Long, j As Long, cellValue As Double
    Dim maxAssociationThreshold As Double, highThreshold As Double
    
    maxAssociationThreshold = Associations_Total - 6
    highThreshold = Associations_Total - 1
    
    Dim targetTables As Variant
    targetTables = Array(pptTableLeft, pptTableRight)
    
    For Each pptTable In targetTables
        For i = 3 To pptTable.Rows.count
            For j = 3 To 4 ' Only column 3 and 4
                On Error Resume Next
                cellValue = CDbl(pptTable.cell(i, j).shape.TextFrame.textRange.text)
                On Error GoTo 0
                
                With pptTable.cell(i, j).shape.Fill
                    .Solid ' Ensure only solid fill is applied
                End With
                
                If cellValue > highThreshold Then
                    pptTable.cell(i, j).shape.Fill.ForeColor.RGB = RGB(153, 50, 75) ' #99324B
                    pptTable.cell(i, j).shape.TextFrame.textRange.Font.color = RGB(255, 255, 255) ' White
                ElseIf cellValue > maxAssociationThreshold Then
                    pptTable.cell(i, j).shape.Fill.ForeColor.RGB = RGB(194, 132, 147) ' #C28493
                    pptTable.cell(i, j).shape.TextFrame.textRange.Font.color = RGB(17, 21, 66) ' 17,21,66
                ElseIf cellValue < 6 Then
                    pptTable.cell(i, j).shape.Fill.ForeColor.RGB = RGB(51, 161, 154) ' #33A19A
                    pptTable.cell(i, j).shape.TextFrame.textRange.Font.color = RGB(255, 255, 255) ' White
                ElseIf cellValue < 11 And cellValue > 5 Then
                    pptTable.cell(i, j).shape.Fill.ForeColor.RGB = RGB(153, 208, 204) ' #99D0CC
                    pptTable.cell(i, j).shape.TextFrame.textRange.Font.color = RGB(17, 21, 66) ' 17,21,66
                End If
            Next j
        Next i
    Next pptTable

    Debug.Print "Color formatting applied to LEFTIE and RIGHTIE."
End Sub

