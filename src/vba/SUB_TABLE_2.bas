Attribute VB_Name = "SUB_TABLE_2"
Sub SUB_TABLE_2()
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
    Dim col4 As String ' Stores extracted column text
    
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

            ' Ensure there are at least 4 columns
            If UBound(Data) >= 3 Then
                ' Extract and clean relevant columns
                Dim col1 As String: col1 = Trim(LCase(Data(0))) ' Convert to lowercase for comparison
                col4 = Trim(Data(3)) ' Keep original case for bullets

                ' Convert col4 to lowercase for filtering
                Dim col4Lower As String: col4Lower = LCase(col4)

                ' Exclude "false", "falskt", empty values, and common misspellings
                If col1 = "stronger" And col4Lower <> "" And _
                   col4Lower <> "false" And col4Lower <> "falskt" And _
                   col4Lower <> "fals" And col4Lower <> "fales" And col4Lower <> "flase" Then

                    ' Ensure no unnecessary blank lines
                    If bulletText <> "" Then bulletText = bulletText & Chr(10)
                    bulletText = bulletText & col4 ' Append cleaned text properly for Mac
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

    ' === Insert Cleaned Text with Bullets in Table Cell (Row 3, Column 4) ===
    With targetTable.cell(3, 4).shape.TextFrame.textRange
        If bulletText <> "" Then
            .text = bulletText ' Insert cleaned text
            .ParagraphFormat.Bullet.visible = msoTrue ' Keep bullets
        Else
            .text = "No valid data found."
        End If
    End With

End Sub


