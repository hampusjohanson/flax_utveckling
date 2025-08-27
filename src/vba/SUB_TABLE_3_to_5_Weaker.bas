Attribute VB_Name = "SUB_TABLE_3_to_5_Weaker"
' === Common Function for "Weaker" Rows ===
Sub InsertDataIntoTarget_Weaker(csvColumnIndex As Integer, targetColumn As Integer)
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

        ' Process only rows 1161-1190
        If rowIndex >= 1161 And rowIndex <= 1190 Then
            Data = Split(line, ";") ' Use ; as delimiter

            ' Ensure there are enough columns
            If UBound(Data) >= csvColumnIndex Then
                ' Extract and clean relevant columns
                Dim col1 As String: col1 = Trim(LCase(Data(0))) ' Convert to lowercase for comparison
                colX = Trim(Data(csvColumnIndex)) ' Keep original case for bullets

                ' Convert colX to lowercase for filtering
                Dim colXLower As String: colXLower = LCase(colX)

                ' Exclude "false", "falskt", and empty values
                If col1 = "weaker" And colXLower <> "" And _
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

    ' === Insert Cleaned Text with Bullets in Table Cell (Row 5, Target Column) ===
    With targetTable.cell(5, targetColumn).shape.TextFrame.textRange
        If bulletText <> "" Then
            .text = bulletText ' Insert cleaned text
            .ParagraphFormat.Bullet.visible = msoTrue ' Keep bullets
        Else
            .text = "No valid data found."
        End If
    End With

End Sub


' === NEW MACROS FOR "WEAKER" DATA ===

Sub SUB_TABLE_3_Weaker()
    InsertDataIntoTarget_Weaker 2, 3 ' Column D (Index X in VBA, Insert into Column X in Table)
End Sub

Sub SUB_TABLE_4_Weaker()
    InsertDataIntoTarget_Weaker 3, 4 ' Column E (Index X in VBA, Insert into Column X in Table)
End Sub

Sub SUB_TABLE_5_Weaker()
    InsertDataIntoTarget_Weaker 4, 5 ' Column F (Index X in VBA, Insert into Column X in Table)
End Sub
Sub SUB_TABLE_6_Weaker()
    InsertDataIntoTarget_Weaker 5, 6 ' Column F (Index X in VBA, Insert into Column X in Table)
End Sub
Sub SUB_TABLE_7_Weaker()
    InsertDataIntoTarget_Weaker 6, 7 ' Column F (Index X in VBA, Insert into Column X in Table)
End Sub

