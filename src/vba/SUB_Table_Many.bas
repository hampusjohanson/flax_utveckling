Attribute VB_Name = "SUB_Table_Many"
' === Common Function to Read and Insert Data ===
Sub InsertDataIntoTarget(csvColumnIndex As Integer, targetColumn As Integer)
    ' Variables
    Dim pptSlide As slide
    Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim Data() As String
    Dim rowIndex As Integer
    Dim tableShape As shape
    Dim targetTable As table
    Dim operatingSystem As String
    Dim userName As String
    Dim bulletText As String
    Dim colX As String ' Stores extracted column text

    ' === Detect Operating System & Username ===
    userName = Environ("USER")
    operatingSystem = Application.operatingSystem

    If InStr(operatingSystem, "Macintosh") > 0 Then
        filePath = "/Users/" & userName & "/Desktop/exported_data_semi.csv"
    Else
        filePath = "C:\Local\exported_data_semi.csv"
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    ' === Read CSV File ===
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    rowIndex = 0
    bulletText = "" ' Initialize empty string

    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        rowIndex = rowIndex + 1

        ' Process only rows 1131-1160
        If rowIndex >= 1131 And rowIndex <= 1160 Then
            Data = Split(line, ";") ' Use ; as delimiter

            ' Ensure there are enough columns
            If UBound(Data) >= csvColumnIndex Then
                ' Extract and clean relevant columns
                Dim col1 As String: col1 = Trim(LCase(Data(0))) ' Convert to lowercase for comparison
                colX = Trim(Data(csvColumnIndex)) ' Keep original case for bullets

                ' Convert colX to lowercase for filtering
                Dim colXLower As String: colXLower = LCase(colX)

                ' Exclude "false", "falskt", and empty values
                If col1 = "stronger" And colXLower <> "" And _
                   colXLower <> "false" And colXLower <> "falskt" And _
                   colXLower <> "fals" And colXLower <> "fales" And colXLower <> "flase" Then

                    ' Ensure no unnecessary blank lines
                    If bulletText <> "" Then bulletText = bulletText & Chr(10)
                    bulletText = bulletText & colX ' Append cleaned text properly for Mac
                End If
            End If
        End If
    Loop
    Close fileNumber

    ' === Find TARGET Table ===
    Set pptSlide = ActiveWindow.View.slide
    Set tableShape = Nothing

    For Each tableShape In pptSlide.Shapes
        If tableShape.HasTable Then
            If tableShape.Name = "TARGET" Then
                Set targetTable = tableShape.table
                Exit For
            End If
        End If
    Next tableShape

    ' If TARGET table is not found
    If targetTable Is Nothing Then
        MsgBox "Table 'TARGET' not found on the slide.", vbExclamation
        Exit Sub
    End If

    ' === Insert Cleaned Text with Bullets in Table Cell (Row 3, Target Column) ===
    With targetTable.cell(3, targetColumn).shape.TextFrame.textRange
        If bulletText <> "" Then
            .text = bulletText ' Insert cleaned text
            .ParagraphFormat.Bullet.visible = msoTrue ' Keep bullets
        Else
            .text = "No valid data found."
        End If
    End With

End Sub


' === INDIVIDUAL MACROS CALLING THE FUNCTION ===

Sub SUB_TABLE_3()
    InsertDataIntoTarget 3, 4 ' Column D (Index 3 in VBA, Insert into Column 5 in Table)
End Sub

Sub SUB_TABLE_4()
    InsertDataIntoTarget 4, 5 ' Column E (Index 4 in VBA, Insert into Column 6 in Table)
End Sub

Sub SUB_TABLE_5()
    InsertDataIntoTarget 5, 6 ' Column F (Index 5 in VBA, Insert into Column 7 in Table)
End Sub
Sub SUB_TABLE_6()
    InsertDataIntoTarget 6, 7 ' Column F (Index 5 in VBA, Insert into Column 7 in Table)
End Sub
Sub SUB_TABLE_7()
    InsertDataIntoTarget 6, 8 ' Column F (Index 5 in VBA, Insert into Column 7 in Table)
End Sub
Sub SUB_TABLE_8()
    InsertDataIntoTarget 2, 3 ' Column D (Index 3 in VBA, Insert into Column 5 in Table)
End Sub
Sub SUB_TABLE_9()
    InsertDataIntoTarget 6, 7 ' Column D (Index 3 in VBA, Insert into Column 5 in Table)
End Sub

