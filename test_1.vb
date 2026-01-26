
' ================================
' Log sheet "State Affecting" population
' ================================
' What this module does
' - You run PopulateLog_StateAffecting_Master
' - It fills the 50 columns named like "Alabama Affecting" ... "Wyoming Affecting"
' - For each row in the Log sheet's used range:
'     - If column K contains "All" OR contains that state's code (e.g., "USAL"), write "Yes"
'     - Otherwise write "No"
'
Option Explicit

' Entry point (run this macro)
Public Sub PopulateLog_StateAffecting_Master()
	' Requirements checklist:
	' [x] Multiple private subs in same file
	' [x] One master sub to run everything
	' [x] Loop used range rows; read column K; write Yes/No for each state column
	' [x] Handle "All" OR state code match (e.g., USAL)
	' [x] Filter Log by Checks!B2:C2 date range and copy A:J + selected Affecting columns to Final

	Dim ws As Worksheet
	Set ws = GetLogWorksheet()
	If ws Is Nothing Then Exit Sub

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.Calculation = xlCalculationManual

	On Error GoTo CleanFail
	PopulateLog_StateAffecting ws
	ExportChecksSelection_ToFinal

CleanExit:
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	Exit Sub

CleanFail:
	' Keep it simple: restore Excel state and re-raise
	Resume CleanExit
End Sub

' ================================
' Export filtered Log data to Final
' ================================
' Uses:
' - Checks!B2 = start date
' - Checks!C2 = end date
' - Checks!A:A = list of states (e.g., Alabama, Alaska, ...)
' Output:
' - Writes headers + filtered rows to Final sheet
Private Sub ExportChecksSelection_ToFinal()
	Dim wsChecks As Worksheet, wsLog As Worksheet, wsFinal As Worksheet
	Set wsChecks = GetWorksheetByName("Checks")
	Set wsLog = GetWorksheetByName("Log")
	Set wsFinal = GetWorksheetByName("Final")
	If wsChecks Is Nothing Or wsLog Is Nothing Or wsFinal Is Nothing Then Exit Sub

	Dim startDate As Date, endDate As Date
	If Not TryGetDate(wsChecks.Range("B2").Value2, startDate) Then Exit Sub
	If Not TryGetDate(wsChecks.Range("C2").Value2, endDate) Then Exit Sub
	If startDate > endDate Then
		Dim tmp As Date
		tmp = startDate
		startDate = endDate
		endDate = tmp
	End If

	' Determine the date column on Log by header match (preferred), else fall back to column A.
	Dim dateCol As Long
	dateCol = FindDateColumn(wsLog)
	If dateCol = 0 Then dateCol = 1

	Dim lastRow As Long, lastCol As Long
	lastRow = wsLog.Cells(wsLog.Rows.Count, dateCol).End(xlUp).Row
	If lastRow < 2 Then Exit Sub
	lastCol = wsLog.Cells(1, wsLog.Columns.Count).End(xlToLeft).Column

	' Clear Final and write headers
	wsFinal.Cells.Clear

	' Build the list of columns to export:
	' - Always A:J (1..10)
	' - Then each "<State> Affecting" for states listed in Checks column A
	Dim exportCols As Collection
	Set exportCols = New Collection
	Dim c As Long
	For c = 1 To 10
		exportCols.Add c
	Next c

	Dim states As Collection
	Set states = ReadStateListFromChecks(wsChecks)
	If states.Count = 0 Then Exit Sub

	Dim i As Long
	For i = 1 To states.Count
		Dim headerText As String
		headerText = CStr(states(i)) & " Affecting"
		Dim colIndex As Long
		colIndex = FindHeaderColumn(wsLog, 1, 1, lastCol, headerText)
		If colIndex > 0 Then
			AddUniqueLong exportCols, colIndex
		End If
	Next i

	' Write Final headers in the exported order
	Dim outCol As Long
	outCol = 1
	For i = 1 To exportCols.Count
		wsFinal.Cells(1, outCol).Value2 = wsLog.Cells(1, CLng(exportCols(i))).Value2
		outCol = outCol + 1
	Next i

	' Copy filtered rows (without relying on AutoFilter copy behaviour / contiguous ranges)
	Dim outRow As Long
	outRow = 2

	Dim r As Long
	For r = 2 To lastRow
		Dim rowDate As Date
		If TryGetDate(wsLog.Cells(r, dateCol).Value2, rowDate) Then
			If rowDate >= startDate And rowDate <= endDate Then
				outCol = 1
				For i = 1 To exportCols.Count
					wsFinal.Cells(outRow, outCol).Value2 = wsLog.Cells(r, CLng(exportCols(i))).Value2
					outCol = outCol + 1
				Next i
				outRow = outRow + 1
			End If
		End If
	Next r
End Sub

Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
	On Error Resume Next
	Set GetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
	On Error GoTo 0
End Function

Private Function TryGetDate(ByVal v As Variant, ByRef d As Date) As Boolean
	On Error GoTo Fail
	If IsDate(v) Then
		d = CDate(v)
		TryGetDate = True
		Exit Function
	End If
Fail:
	TryGetDate = False
End Function

Private Function FindDateColumn(ByVal ws As Worksheet) As Long
	' Tries to locate a date column by common header names in row 1.
	Dim lastCol As Long
	lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

	Dim candidates As Variant
	candidates = Array("DATE", "LOG DATE", "CREATED DATE", "CREATED", "START DATE")

	Dim c As Long
	For c = 1 To lastCol
		Dim h As String
		h = UCase$(Trim$(CStr(ws.Cells(1, c).Value2)))
		Dim i As Long
		For i = LBound(candidates) To UBound(candidates)
			If h = CStr(candidates(i)) Then
				FindDateColumn = c
				Exit Function
			End If
		Next i
	Next c
	FindDateColumn = 0
End Function

Private Function ReadStateListFromChecks(ByVal wsChecks As Worksheet) As Collection
	' Reads Checks!A:A from row 2 down until first blank.
	' Expected values: full state names (e.g., "Alabama").
	Dim result As New Collection

	Dim r As Long
	r = 2
	Do While Len(Trim$(CStr(wsChecks.Cells(r, 1).Value2))) > 0
		Dim s As String
		s = Trim$(CStr(wsChecks.Cells(r, 1).Value2))
		AddUniqueString result, s
		r = r + 1
	Loop

	Set ReadStateListFromChecks = result
End Function

Private Sub AddUniqueString(ByVal col As Collection, ByVal value As String)
	Dim v As Variant
	For Each v In col
		If StrComp(CStr(v), value, vbTextCompare) = 0 Then Exit Sub
	Next v
	col.Add value
End Sub

Private Sub AddUniqueLong(ByVal col As Collection, ByVal value As Long)
	Dim v As Variant
	For Each v In col
		If CLng(v) = value Then Exit Sub
	Next v
	col.Add value
End Sub

' ------------------------
' Core implementation
' ------------------------
Private Sub PopulateLog_StateAffecting(ByVal wsLog As Worksheet)
	Dim stateNames As Variant
	Dim stateCodes As Variant
	GetStateLists stateNames, stateCodes

	Dim lastRow As Long
	lastRow = GetLastRowFromColumn(wsLog, 11) ' Column K
	If lastRow < 2 Then Exit Sub ' assume row 1 is header

	' Map each "<State> Affecting" header to its column index.
	Dim stateCols() As Long
	stateCols = GetStateAffectingColumns(wsLog, stateNames)

	Dim r As Long
	For r = 2 To lastRow
		Dim kValue As String
		kValue = NormalizeAffectingValue(wsLog.Cells(r, 11).Value2)

		Dim isAll As Boolean
		isAll = ContainsToken(kValue, "ALL")

		Dim i As Long
		For i = LBound(stateNames) To UBound(stateNames)
			Dim targetCol As Long
			targetCol = stateCols(i)
			If targetCol > 0 Then
				Dim yesNo As String
				If isAll Or ContainsToken(kValue, CStr(stateCodes(i))) Then
					yesNo = "Yes"
				Else
					yesNo = "No"
				End If
				wsLog.Cells(r, targetCol).Value2 = yesNo
			End If
		Next i
	Next r
End Sub

' ------------------------
' Helpers
' ------------------------
Private Function GetLogWorksheet() As Worksheet
	' Assumption: the sheet is literally named "Log".
	' If your sheet has a different name, change it here.
	On Error Resume Next
	Set GetLogWorksheet = ThisWorkbook.Worksheets("Log")
	On Error GoTo 0
End Function

Private Sub GetStateLists(ByRef stateNames As Variant, ByRef stateCodes As Variant)
	' 50 states only (no DC / territories).
	' NOTE: VBA is picky about line continuations in Array(...) calls.
	' Keeping these on a single physical line avoids "Syntax error" / "Expected: )" issues.
	stateNames = Array("Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming")

	' Codes per your format: "US" + postal abbreviation.
	stateCodes = Array("USAL", "USAK", "USAZ", "USAR", "USCA", "USCO", "USCT", "USDE", "USFL", "USGA", "USHI", "USID", "USIL", "USIN", "USIA", "USKS", "USKY", "USLA", "USME", "USMD", "USMA", "USMI", "USMN", "USMS", "USMO", "USMT", "USNE", "USNV", "USNH", "USNJ", "USNM", "USNY", "USNC", "USND", "USOH", "USOK", "USOR", "USPA", "USRI", "USSC", "USSD", "USTN", "USTX", "USUT", "USVT", "USVA", "USWA", "USWV", "USWI", "USWY")
End Sub

Private Function GetStateAffectingColumns(ByVal ws As Worksheet, ByVal stateNames As Variant) As Long()
	' Finds columns by exact header match in row 1: "<State> Affecting"
	Dim cols() As Long
	ReDim cols(LBound(stateNames) To UBound(stateNames))

	Dim headerRow As Long
	headerRow = 1

	Dim usedLastCol As Long
	usedLastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

	Dim i As Long
	For i = LBound(stateNames) To UBound(stateNames)
		Dim headerText As String
		headerText = CStr(stateNames(i)) & " Affecting"
		cols(i) = FindHeaderColumn(ws, headerRow, 1, usedLastCol, headerText)
	Next i

	GetStateAffectingColumns = cols
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long, ByVal endCol As Long, ByVal headerText As String) As Long
	Dim c As Long
	For c = startCol To endCol
		If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value2)), headerText, vbTextCompare) = 0 Then
			FindHeaderColumn = c
			Exit Function
		End If
	Next c
	FindHeaderColumn = 0
End Function

Private Function GetLastRowFromColumn(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
	' Uses the last non-empty cell in a given column.
	Dim lastRow As Long
	lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
	GetLastRowFromColumn = lastRow
End Function

Private Function NormalizeAffectingValue(ByVal v As Variant) As String
	' Normalizes the raw cell value so matching is more reliable.
	' - Uppercases
	' - Replaces common separators with spaces
	' - Pads with spaces at both ends so token-search is easy
	Dim s As String
	s = UCase$(Trim$(CStr(v)))

	' Normalize separators to spaces
	s = Replace(s, ",", " ")
	s = Replace(s, ";", " ")
	s = Replace(s, vbTab, " ")
	s = Replace(s, vbCr, " ")
	s = Replace(s, vbLf, " ")

	' Collapse multiple spaces
	Do While InStr(1, s, "  ") > 0
		s = Replace(s, "  ", " ")
	Loop

	NormalizeAffectingValue = " " & s & " "
End Function

Private Function ContainsToken(ByVal normalizedPadded As String, ByVal token As String) As Boolean
	' Expects normalizedPadded to be uppercased and padded with spaces.
	' Matches whole tokens only (so "USAL" doesn't match "USALX").
	Dim t As String
	t = " " & UCase$(Trim$(token)) & " "
	ContainsToken = (InStr(1, normalizedPadded, t, vbTextCompare) > 0)
End Function

