Private Sub CheckA4MatchesA1()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim a1 As String, a4 As String
    Dim firstClose As Long, nextOpen As Long
    Dim extracted As String, firstLine As String

    a1 = CStr(ws.Range("A1").Value)
    a4 = CStr(ws.Range("A4").Value)

    ' Find first ")" in A1
    firstClose = InStr(1, a1, ")")
    If firstClose = 0 Then
        MsgBox "A1 does not contain a closing parenthesis.", vbExclamation
        Exit Sub
    End If

    ' Find next "(" after that
    nextOpen = InStr(firstClose + 1, a1, "(")
    If nextOpen = 0 Then
        MsgBox "A1 does not contain a second opening parenthesis.", vbExclamation
        Exit Sub
    End If

    ' Extract the text between ) and (
    extracted = Trim(Mid$(a1, firstClose + 1, nextOpen - firstClose - 1))

    ' Extract the first line of A4
    If InStr(a4, vbLf) > 0 Then
        firstLine = Trim(Left$(a4, InStr(a4, vbLf) - 1))
    Else
        firstLine = Trim(a4)
    End If

    ' Compare
    If StrComp(extracted, firstLine, vbBinaryCompare) = 0 Then
        MsgBox "A4 matches the value in A1.", vbInformation
    Else
        MsgBox "A4 does NOT match the value in A1." & vbCrLf & _
               "Expected: " & extracted & vbCrLf & _
               "Found: " & firstLine, vbCritical
    End If

End Sub
