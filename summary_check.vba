Option Explicit

'==========================
'  HEADER LOOKUP FUNCTION
'==========================
Function GetColumnByHeader(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Range
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For Each c In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If StrComp(Trim(c.Value), headerName, vbBinaryCompare) = 0 Then
            GetColumnByHeader = c.Column
            Exit Function
        End If
    Next c

    GetColumnByHeader = 0
End Function

'==========================
'  LOAD SUMMARY DICTIONARY
'==========================
Function LoadSummaryDict() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Summary")

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long, key As String, val As String

    For i = 2 To lastRow
        key = ws.Cells(i, "A").Value
        val = ws.Cells(i, "B").Value

        If Len(key) > 0 Then
            dict(key) = val
        End If
    Next i

    Set LoadSummaryDict = dict
End Function

'==========================
'  HIGHLIGHT MATCHES
'==========================
Sub HighlightExactMatches(rng As Range, dict As Object)
    Dim cell As Range, key As Variant
    Dim pos As Long

    For Each cell In rng.Cells
        If Len(cell.Value) > 0 Then
            For Each key In dict.Keys
                pos = InStrB(1, cell.Value, key, vbBinaryCompare)
                If pos > 0 Then
                    ' Convert byte position to character position
                    pos = (pos + 1) \ 2
                    cell.Characters(pos, Len(key)).Font.Color = vbRed
                End If
            Next key
        End If
    Next cell
End Sub

'==========================
'  REPLACE MATCHES
'==========================
Sub ReplaceExactMatches(rng As Range, dict As Object)
    Dim cell As Range, key As Variant
    Dim pos As Long, replacement As String

    For Each cell In rng.Cells
        If Len(cell.Value) > 0 Then
            For Each key In dict.Keys
                pos = InStrB(1, cell.Value, key, vbBinaryCompare)
                If pos > 0 Then
                    replacement = dict(key)
                    If replacement = "" Then
                        cell.Value = Replace(cell.Value, key, "", 1, 1, vbBinaryCompare)
                    Else
                        cell.Value = Replace(cell.Value, key, replacement, 1, 1, vbBinaryCompare)
                    End If
                End If
            Next key
        End If
    Next cell
End Sub

'==========================
'  MAIN CONTROLLER MACRO
'==========================
Sub ProcessSoC()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Checks")

    Dim socCol As Long, updatedCol As Long
    socCol = GetColumnByHeader(ws, "SoC")
    updatedCol = GetColumnByHeader(ws, "Updated SoC")

    If socCol = 0 Or updatedCol = 0 Then
        MsgBox "Could not find 'SoC' or 'Updated SoC' column headers.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, socCol).End(xlUp).Row

    Dim socRange As Range, updatedRange As Range
    Set socRange = ws.Range(ws.Cells(2, socCol), ws.Cells(lastRow, socCol))
    Set updatedRange = ws.Range(ws.Cells(2, updatedCol), ws.Cells(lastRow, updatedCol))

    Dim dict As Object
    Set dict = LoadSummaryDict()

    ' Step 1: Highlight matches in SoC
    HighlightExactMatches socRange, dict

    ' Step 2: Copy SoC → Updated SoC
    updatedRange.Value = socRange.Value

    ' Step 3: Replace matches in Updated SoC
    ReplaceExactMatches updatedRange, dict

    MsgBox "SoC processing complete.", vbInformation
End Sub

'==========================
'  Copy / Paste from the Updated SoC to the SoC column
'==========================

Sub CopyUpdatedSoCToSoC()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Checks")

    Dim socCol As Long, updatedCol As Long
    socCol = GetColumnByHeader(ws, "SoC")
    updatedCol = GetColumnByHeader(ws, "Updated SoC")

    If socCol = 0 Or updatedCol = 0 Then
        MsgBox "Could not find 'SoC' or 'Updated SoC' column headers.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, updatedCol).End(xlUp).Row

    ws.Range(ws.Cells(2, socCol), ws.Cells(lastRow, socCol)).Value = _
        ws.Range(ws.Cells(2, updatedCol), ws.Cells(lastRow, updatedCol)).Value

End Sub

