
Option Explicit

' Scans the active worksheet for columns named:
'   - Ref
'   - Cat
'   - Person
' For each unique Ref that has multiple distinct Cat values, writes the offending
' rows to a clean table on the Review sheet.
Public Sub BuildReview_RefCatConflicts()
	Dim wsData As Worksheet
	Set wsData = ActiveSheet
	If wsData Is Nothing Then Exit Sub
	If StrComp(wsData.Name, "Review", vbTextCompare) = 0 Then
		MsgBox "Please run this macro from your data sheet (not the Review tab).", vbExclamation, "Ref/Cat Review"
		Exit Sub
	End If

	Dim refCol As Long, catCol As Long, personCol As Long, crqCol As Long
	refCol = FindHeaderColumn(wsData, 1, "Ref")
	catCol = FindHeaderColumn(wsData, 1, "Cat")
	personCol = FindHeaderColumn(wsData, 1, "Person")
	crqCol = FindHeaderColumn(wsData, 1, "CRQ")

	If refCol = 0 Or catCol = 0 Or personCol = 0 Or crqCol = 0 Then
		Dim missing As String
		missing = "Missing required header(s):" & vbCrLf
		If refCol = 0 Then missing = missing & "- Ref" & vbCrLf
		If catCol = 0 Then missing = missing & "- Cat" & vbCrLf
		If personCol = 0 Then missing = missing & "- Person" & vbCrLf
		If crqCol = 0 Then missing = missing & "- CRQ" & vbCrLf
		MsgBox missing, vbExclamation, "Ref/Cat Review"
		Exit Sub
	End If

	Dim lastRow As Long
	lastRow = wsData.Cells(wsData.Rows.Count, refCol).End(xlUp).Row
	If lastRow < 2 Then
		MsgBox "No data rows found under the Ref header.", vbInformation, "Ref/Cat Review"
		Exit Sub
	End If

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.Calculation = xlCalculationManual

	On Error GoTo CleanFail

	Dim catsPerRef As Object
	Set catsPerRef = CreateObject("Scripting.Dictionary") ' Ref -> Dictionary(Cat -> True)
	Dim rowsPerRef As Object
	Set rowsPerRef = CreateObject("Scripting.Dictionary") ' Ref -> Collection of row payloads

	Dim r As Long
	For r = 2 To lastRow
		Dim refValue As String
		refValue = Trim$(CStr(wsData.Cells(r, refCol).Value2))
		If Len(refValue) = 0 Then GoTo NextRow

		Dim catValue As String
		catValue = Trim$(CStr(wsData.Cells(r, catCol).Value2))
		Dim personValue As String
		personValue = Trim$(CStr(wsData.Cells(r, personCol).Value2))
		Dim crqValue As String
		crqValue = Trim$(CStr(wsData.Cells(r, crqCol).Value2))

		If Not catsPerRef.Exists(refValue) Then
			Dim d As Object
			Set d = CreateObject("Scripting.Dictionary")
			catsPerRef.Add refValue, d
			Dim c As Collection
			Set c = New Collection
			rowsPerRef.Add refValue, c
		End If

		Dim catsDict As Object
		Set catsDict = catsPerRef(refValue)
		If Not catsDict.Exists(catValue) Then catsDict.Add catValue, True

		Dim payload(1 To 4) As Variant
		payload(1) = catValue
		payload(2) = personValue
		payload(3) = crqValue
		payload(4) = r
		rowsPerRef(refValue).Add payload

NextRow:
	Next r

	Dim wsReview As Worksheet
	Set wsReview = GetOrCreateWorksheet("Review")
	If wsReview Is Nothing Then GoTo CleanFail

	ResetReviewSheet wsReview

	' Headers
	wsReview.Range("A1").Value2 = "Ref"
	wsReview.Range("B1").Value2 = "Cat"
	wsReview.Range("C1").Value2 = "Person"
	wsReview.Range("D1").Value2 = "CRQ"
	wsReview.Range("E1").Value2 = "SourceRow"
	wsReview.Range("F1").Value2 = "CatsForRef"
	wsReview.Range("G1").Value2 = "SourceSheet"

	Dim outRow As Long
	outRow = 2

	Dim refKey As Variant
	For Each refKey In catsPerRef.Keys
		Dim perRefCats As Object
		Set perRefCats = catsPerRef(refKey)
		If perRefCats.Count > 1 Then
			Dim catsText As String
			catsText = JoinDictionaryKeys(perRefCats, " | ")

			Dim item As Variant
			For Each item In rowsPerRef(refKey)
				wsReview.Cells(outRow, 1).Value2 = CStr(refKey)
				wsReview.Cells(outRow, 2).Value2 = CStr(item(1))
				wsReview.Cells(outRow, 3).Value2 = CStr(item(2))
				wsReview.Cells(outRow, 4).Value2 = CStr(item(3))
				wsReview.Cells(outRow, 5).Value2 = CLng(item(4))
				wsReview.Cells(outRow, 6).Value2 = catsText
				wsReview.Cells(outRow, 7).Value2 = wsData.Name
				outRow = outRow + 1
			Next item
		End If
	Next refKey

	If outRow = 2 Then
		wsReview.Range("A3").Value2 = "No Ref values found with multiple Cat values."
	End If

	CreateReviewTable wsReview, outRow - 1

	wsReview.Range("A1").Select
	wsReview.Activate

CleanExit:
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	Exit Sub

CleanFail:
	Resume CleanExit
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
	Dim lastCol As Long
	lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
	If lastCol < 1 Then
		FindHeaderColumn = 0
		Exit Function
	End If

	Dim target As String
	target = UCase$(Trim$(headerText))

	Dim c As Long
	For c = 1 To lastCol
		Dim h As String
		h = UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value2)))
		If h = target Then
			FindHeaderColumn = c
			Exit Function
		End If
	Next c

	FindHeaderColumn = 0
End Function

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
	On Error Resume Next
	Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
	On Error GoTo 0

	If GetOrCreateWorksheet Is Nothing Then
		On Error Resume Next
		Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
		If Not GetOrCreateWorksheet Is Nothing Then GetOrCreateWorksheet.Name = sheetName
		On Error GoTo 0
	End If
End Function

Private Sub ResetReviewSheet(ByVal wsReview As Worksheet)
	' Clear content and remove existing tables for a clean rebuild.
	Dim lo As ListObject
	For Each lo In wsReview.ListObjects
		lo.Unlist
	Next lo
	wsReview.Cells.Clear

	wsReview.Range("A1").Font.Bold = True
	wsReview.Range("A1").EntireRow.Font.Bold = True
	wsReview.Range("A1").EntireRow.WrapText = False
End Sub

Private Sub CreateReviewTable(ByVal wsReview As Worksheet, ByVal lastDataRow As Long)
	Dim lastCol As Long
	lastCol = 7

	Dim rng As Range
	If lastDataRow < 1 Then lastDataRow = 1
	Set rng = wsReview.Range(wsReview.Cells(1, 1), wsReview.Cells(Application.WorksheetFunction.Max(1, lastDataRow), lastCol))

	Dim lo As ListObject
	Set lo = wsReview.ListObjects.Add(xlSrcRange, rng, , xlYes)
	lo.Name = "tblRefCatReview"
	lo.TableStyle = "TableStyleMedium2"

	wsReview.Columns("A:G").EntireColumn.AutoFit
	wsReview.Range("A1").EntireRow.AutoFilter
	wsReview.Range("A1").EntireRow.HorizontalAlignment = xlLeft
End Sub

Private Function JoinDictionaryKeys(ByVal dict As Object, ByVal separator As String) As String
	Dim result As String
	Dim k As Variant
	For Each k In dict.Keys
		If Len(result) > 0 Then result = result & separator
		result = result & CStr(k)
	Next k
	JoinDictionaryKeys = result
End Function

