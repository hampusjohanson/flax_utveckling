Attribute VB_Name = "Module19"
Option Explicit

#If Win64 Or Win32 Then
    ' Declare GetAsyncKeyState for Windows
    #If VBA7 Then
        Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    #Else
        Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    #End If
#End If

Sub labeling_move_with_arrows()
    Dim moveX As Single, moveY As Single
    Dim os As String

    ' Detect OS
    If Environ("OS") Like "*Windows*" Then
        os = "Windows"
    Else
        os = "Mac"
    End If

    ' Ensure a label has been stored
    If LabelToMove Is Nothing Then
        MsgBox "No label has been stored for movement. Run the previous macro first.", vbExclamation
        Exit Sub
    End If

    ' **Select the label in PowerPoint**
    LabelToMove.Parent.Select
    LabelToMove.Select

    ' Instructions
    MsgBox "Hold 'B' and use arrow keys (Windows) or W/A/S/D (Mac) to move the label. Press Enter or Esc (Windows) or 'Q' (Mac) to exit.", vbInformation

    ' **Windows: Use real-time movement**
    If os = "Windows" Then
        Do
            ' **Only move if "B" is held**
            If GetAsyncKeyState(Asc("B")) <> 0 Then
                If GetAsyncKeyState(37) <> 0 Then LabelToMove.left = LabelToMove.left - 2 ' Left
                If GetAsyncKeyState(39) <> 0 Then LabelToMove.left = LabelToMove.left + 2 ' Right
                If GetAsyncKeyState(38) <> 0 Then LabelToMove.Top = LabelToMove.Top - 2 ' Up
                If GetAsyncKeyState(40) <> 0 Then LabelToMove.Top = LabelToMove.Top + 2 ' Down
            End If

            ' **Exit when Enter or Esc is pressed**
            If GetAsyncKeyState(13) <> 0 Or GetAsyncKeyState(27) <> 0 Then Exit Do

            DoEvents ' Keeps it responsive
        Loop

    ' **Mac: Use manual key input**
    Else
        Do
            If MacDetectKeyPress("b") Then
                Select Case MacDetectKeyPress("")
                    Case "w": LabelToMove.Top = LabelToMove.Top - 2 ' Up
                    Case "s": LabelToMove.Top = LabelToMove.Top + 2 ' Down
                    Case "a": LabelToMove.left = LabelToMove.left - 2 ' Left
                    Case "d": LabelToMove.left = LabelToMove.left + 2 ' Right
                End Select
            End If
            
            ' **Exit on "Q"**
            If MacDetectKeyPress("q") Then Exit Do

            DoEvents
        Loop
    End If

    ' **Confirm movement**
    Debug.Print "Final position of label [" & LabelToMove.text & "] - X: " & LabelToMove.left & ", Y: " & LabelToMove.Top
End Sub

' **Mac Key Detection (Lightweight)**
Function MacDetectKeyPress(key As String) As Boolean
    Dim script As String
    script = "tell application ""System Events"" to return (keystroke """ & key & """)"
    On Error Resume Next
    MacDetectKeyPress = (MacScript(script) <> "")
    On Error GoTo 0
End Function

