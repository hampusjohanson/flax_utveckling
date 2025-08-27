Attribute VB_Name = "Text_Input_1"
Sub Text_Input_1()
    Dim pptSlide As slide
    Dim s As shape
    Dim tbl As table
    Dim rowIndex As Integer
    Dim textValue As String
    Dim occurrences() As Variant
    Dim tblCount As Integer
    Dim hasBold As Boolean
    Dim leftMostTable As shape
    Dim txtBox As shape
    Dim textRange As textRange
    Dim boldText As textRange
    Dim rowCount As Integer
    Dim minBoldThreshold As Integer
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim uniqueBoldCount As Integer
    
    ' Set the active slide
    Set pptSlide = ActiveWindow.View.slide
    
    ' Initialize variables
    tblCount = 0
    hasBold = False
    uniqueBoldCount = 0
    Set leftMostTable = Nothing

    ' Initialize occurrences array (2D array: column 1 for text, column 2 for counts)
    ReDim occurrences(1 To 100, 1 To 2) ' Adjust the size as necessary for your data
    Dim uniqueCount As Integer
    uniqueCount = 0

    ' Iterate through all shapes on the slide
    For Each s In pptSlide.Shapes
        If s.HasTable Then
            tblCount = tblCount + 1
            Set tbl = s.table
            
            ' If this is the first table, set leftMostTable
            If leftMostTable Is Nothing Then
                Set leftMostTable = s
            ElseIf s.left < leftMostTable.left Then
                Set leftMostTable = s
            End If
            
            ' Scan column 2 and store only BOLDED text occurrences in the array
            For rowIndex = 2 To tbl.Rows.count
                On Error Resume Next
                textValue = Trim(tbl.cell(rowIndex, 2).shape.TextFrame.textRange.text)
                
                ' Check if the text in the cell is bold
                If tbl.cell(rowIndex, 2).shape.TextFrame.textRange.Font.Bold = msoTrue Then
                    hasBold = True
                    found = False
                    ' Check if text already exists in the occurrences array
                    For i = 1 To uniqueCount
                        If occurrences(i, 1) = textValue Then
                            occurrences(i, 2) = occurrences(i, 2) + 1
                            found = True
                            Exit For
                        End If
                    Next i
                    
                    ' If text is not found, add it to the array with count 1
                    If Not found Then
                        uniqueCount = uniqueCount + 1
                        occurrences(uniqueCount, 1) = textValue
                        occurrences(uniqueCount, 2) = 1 ' Initial count
                    End If
                End If
                On Error GoTo 0
            Next rowIndex
        End If
    Next s

    ' If no bold text is found, exit
    If Not hasBold Then
        Debug.Print "No bold text found. Text box not inserted."
        Exit Sub
    End If

    ' Get the row count for the left-most table
    rowCount = leftMostTable.table.Rows.count

    ' Find the minimum occurrence count among all bold texts
    minBoldThreshold = tblCount ' Default to the total number of tables
    uniqueBoldCount = uniqueCount ' Unique number of bold texts
    For i = 1 To uniqueCount
        If occurrences(i, 2) < minBoldThreshold Then
            minBoldThreshold = occurrences(i, 2)
        End If
    Next i

    ' Create a new text box named "Text_Bold" at the left of the left-most table, 1 cm below it
    Set txtBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                     leftMostTable.left, leftMostTable.Top + leftMostTable.height + 0.25 * 28.35, _
                     400, 50) ' 1 cm below is 28.35 points

    ' Set the name of the text box
    txtBox.Name = "Text_Bold"

    ' Set the text content with rowCount and uniqueBoldCount
    txtBox.TextFrame.textRange.text = "Bold = Top " & rowCount - 1 & " in " & minBoldThreshold & " out of " & tblCount & " tables (" & uniqueBoldCount & " associations)"

    ' Set font size, color, and bold
    txtBox.TextFrame.textRange.Font.size = 10
    txtBox.TextFrame.textRange.Font.Name = "Arial"
    txtBox.TextFrame.textRange.Font.color.RGB = RGB(17, 21, 66)

    ' Make "Bold" part of the text bold
    Set textRange = txtBox.TextFrame.textRange
    Set boldText = textRange.Characters(InStr(1, textRange.text, "Bold"), 4) ' "Bold"
    boldText.Font.Bold = msoTrue

    Debug.Print "Text box 'Text_Bold' created with message: " & txtBox.TextFrame.textRange.text
End Sub


