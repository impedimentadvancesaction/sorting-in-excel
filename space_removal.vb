
Option Explicit

' Cleans text in the active sheet (columns A:F, rows 2..last used row).
' Rules:
' - Removes whitespace at the start/end of the cell (leading/trailing).
' - Converts line breaks (new lines) into single spaces.
' - Preserves spaces between words/values (no collapsing of internal spaces).
' Notes:
' - Only changes cells that contain text (strings). Numbers/dates are left untouched.
' - Errors are ignored.
Public Sub TrimLeadingTrailingWhitespace_AtoF()
	' Work on whichever sheet is currently active.
	Dim ws As Worksheet
	Set ws = ActiveSheet
	If ws Is Nothing Then Exit Sub

	' Determine the last row to process by checking the last used row in each column A..F
	' and taking the maximum. This avoids missing data if one column is longer than others.
	Dim lastRow As Long
	Dim lr As Long
	lastRow = 0
	Dim c As Long
	For c = 1 To 6 ' 1..6 correspond to columns A..F
		lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
		If lr > lastRow Then lastRow = lr
	Next c
	If lastRow < 2 Then Exit Sub

	' Build the target range: A2:F(lastRow).
	Dim rng As Range
	Set rng = ws.Range("A2:F" & lastRow)

	' Speed up large updates and ensure screen updating is restored even if something fails.
	Application.ScreenUpdating = False
	On Error GoTo CleanExit

	' Loop every cell in the target range and clean text values only.
	Dim cell As Range
	For Each cell In rng.Cells
		' Skip Excel error values (e.g., #N/A).
		If Not IsError(cell.Value2) Then
			' Only process strings; leave numbers/dates/booleans unchanged.
			If VarType(cell.Value2) = vbString Then
				Dim s As String
				' Convert any non-breaking spaces to regular spaces.
				' (Non-breaking spaces often come from copied web/email content.)
				s = Replace(CStr(cell.Value2), ChrW$(160), " ")
				' Convert new lines to spaces so multi-line cells become single-line text.
				' vbCrLf = Windows newline, vbCr and vbLf cover other pasted formats.
				s = Replace(s, vbCrLf, " ")
				s = Replace(s, vbCr, " ")
				s = Replace(s, vbLf, " ")
				' Trim removes leading/trailing spaces only (keeps internal spacing).
				cell.Value2 = Trim$(s)
			End If
		End If
	Next cell

CleanExit:
	' Always restore Excel UI updates.
	Application.ScreenUpdating = True
End Sub

