Option Explicit

' Column header names - change these only when your report column titles change
Private Const HEADER_CG As String = "CC"
Private Const HEADER_REF As String = "CRQ"
Private Const HEADER_IMP_DATE As String = "Imp Date"
Private Const HEADER_DESC As String = "Description"
Private Const HEADER_APP_REF As String = "App Ref"
Private Const HEADER_OUTCOME As String = "Results"
Private Const HEADER_CONF As String = "Cert"

' Main subroutine to review a report and log issues
Sub RunReportReview()
    Dim reportPath As String
    Dim reportWb As Workbook
    Dim reportWs As Worksheet
    Dim reviewWs As Worksheet
    Dim keywordsWs As Worksheet
    Dim controlWs As Worksheet
    
    ' Set worksheet references
    On Error Resume Next
    Set controlWs = Worksheets("Review1")
    Set reviewWs = Worksheets("Review2")
    Set keywordsWs = Worksheets("Review3")
    On Error GoTo 0
    
    ' Check if sheets exist
    If controlWs Is Nothing Then
        MsgBox "Review1 sheet does not exist!", vbCritical
        Exit Sub
    End If
    
    If reviewWs Is Nothing Then
        MsgBox "Review2 sheet does not exist!", vbCritical
        Exit Sub
    End If
    
    If keywordsWs Is Nothing Then
        MsgBox "Review3 sheet does not exist!", vbCritical
        Exit Sub
    End If
    
    ' Read file path from Review1 cell A2
    reportPath = Trim(CStr(controlWs.Cells(2, 1).Value))
    
    If reportPath = "" Then
        MsgBox "No file path found in Review1 cell A2!", vbCritical
        Exit Sub
    End If
    
    ' Check if file exists
    If Dir(reportPath) = "" Then
        MsgBox "File not found: " & reportPath, vbCritical
        Exit Sub
    End If
    
    ' Performance optimizations - disable screen updates, calculation, events, and alerts
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean
    Dim originalDisplayAlerts As Boolean
    
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents
    originalDisplayAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Open the report workbook in read-only mode
    On Error GoTo ErrHandler
    Set reportWb = Workbooks.Open(Filename:=reportPath, ReadOnly:=True, UpdateLinks:=False)
    Set reportWs = reportWb.Worksheets(1) ' Assume first sheet contains the data
    
    ' Clear previous review results
    reviewWs.Cells.Clear
    
    ' Initialize review table
    InitializeReviewTable reviewWs
    
    ' Perform all validation checks
    ReviewReportData reportWs, reviewWs, keywordsWs
    
    ' Format the review table
    FormatReviewTable reviewWs
    
    ' Close the report workbook
    reportWb.Close SaveChanges:=False
    
    ' Restore performance settings
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.DisplayAlerts = originalDisplayAlerts
    
    MsgBox "Report review complete! Check Review2 sheet for any issues found.", vbInformation
    Exit Sub
    
ErrHandler:
    ' Restore performance settings even if error occurred
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.DisplayAlerts = originalDisplayAlerts
    
    If Not reportWb Is Nothing Then
        reportWb.Close SaveChanges:=False
    End If
    MsgBox "Error occurred: " & Err.Description, vbCritical
End Sub

' Initialize the review table header
Private Sub InitializeReviewTable(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Issue Type"
        .Cells(1, 2).Value = "Column"
        .Cells(1, 3).Value = "Ref"
        .Cells(1, 4).Value = "Value"
        .Cells(1, 5).Value = "Details"
    End With
End Sub

' Main validation routine
Private Sub ReviewReportData(reportWs As Worksheet, reviewWs As Worksheet, keywordsWs As Worksheet)
    Dim lastRow As Long
    Dim cgCol As Long, refCol As Long, impDateCol As Long, descCol As Long
    Dim appRefCol As Long, outcomeCol As Long, confCol As Long
    Dim i As Long
    Dim nextIssueRow As Long
    Dim forbiddenTerms As Variant
    Dim refValue As String
    
    ' Get column indices by header
    cgCol = GetColumnIndexByHeader(reportWs, HEADER_CG)
    refCol = GetColumnIndexByHeader(reportWs, HEADER_REF)
    impDateCol = GetColumnIndexByHeader(reportWs, HEADER_IMP_DATE)
    descCol = GetColumnIndexByHeader(reportWs, HEADER_DESC)
    appRefCol = GetColumnIndexByHeader(reportWs, HEADER_APP_REF)
    outcomeCol = GetColumnIndexByHeader(reportWs, HEADER_OUTCOME)
    confCol = GetColumnIndexByHeader(reportWs, HEADER_CONF)
    
    ' Log missing headers
    If cgCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_CG, "", "", HEADER_CG & " column header not found"
    If refCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_REF, "", "", HEADER_REF & " column header not found"
    If impDateCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_IMP_DATE, "", "", HEADER_IMP_DATE & " column header not found"
    If descCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_DESC, "", "", HEADER_DESC & " column header not found"
    If appRefCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_APP_REF, "", "", HEADER_APP_REF & " column header not found"
    If outcomeCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_OUTCOME, "", "", HEADER_OUTCOME & " column header not found"
    If confCol = 0 Then LogIssue reviewWs, GetNextIssueRow(reviewWs), "Missing Header", HEADER_CONF, "", "", HEADER_CONF & " column header not found"
    
    ' Load forbidden terms from Review3
    forbiddenTerms = LoadForbiddenTerms(keywordsWs)
    
    ' Get last row of data
    lastRow = reportWs.Cells(reportWs.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        LogIssue reviewWs, GetNextIssueRow(reviewWs), "No Data", "N/A", "", "", "No data rows found in report"
        Exit Sub
    End If
    
    ' Initialize issue row counter
    nextIssueRow = 2
    
    ' Process each data row
    For i = 2 To lastRow
        ' Get Ref value for this row (for review table) - only access cell if column exists
        If refCol > 0 Then
            refValue = Trim(CStr(reportWs.Cells(i, refCol).Value))
        Else
            refValue = ""
        End If
        
        ' Rule 0: Check all cells for whitespace and newline issues
        ValidateCellFormatting reportWs, reviewWs, i, refValue, nextIssueRow
        
        ' Rule 0b: Check for blank cells within used range
        ValidateBlankCells reportWs, reviewWs, i, refValue, nextIssueRow
        
        ' Rule 1: Check CG column
        If cgCol > 0 Then
            ValidateCGColumn reportWs, reviewWs, i, cgCol, refValue, nextIssueRow
        End If
        
        ' Rule 2: Check Ref column
        If refCol > 0 Then
            ValidateRefColumn reportWs, reviewWs, i, refCol, refValue, nextIssueRow
        End If
        
        ' Rule 3: Check Imp Date column
        If impDateCol > 0 Then
            ValidateImpDateColumn reportWs, reviewWs, i, impDateCol, refValue, nextIssueRow
        End If
        
        ' Rule 4: Check Desc column (spell check and forbidden terms)
        If descCol > 0 Then
            ValidateDescColumn reportWs, reviewWs, keywordsWs, i, descCol, refValue, nextIssueRow, forbiddenTerms
        End If
        
        ' Rule 5: Check App Ref vs CG dependency
        If cgCol > 0 And appRefCol > 0 Then
            ValidateAppRefColumn reportWs, reviewWs, i, cgCol, appRefCol, refValue, nextIssueRow
        End If
        
        ' Rule 6: Check Outcome column
        If outcomeCol > 0 Then
            ValidateOutcomeColumn reportWs, reviewWs, i, outcomeCol, refValue, nextIssueRow
        End If
        
        ' Rule 7: Check Conf column
        If confCol > 0 Then
            ValidateConfColumn reportWs, reviewWs, i, confCol, refValue, nextIssueRow
        End If
    Next i
    
    ' Check if any issues were found
    If GetNextIssueRow(reviewWs) = 2 Then
        reviewWs.Cells(2, 1).Value = "No issues found"
        With reviewWs.Cells(2, 1)
            .Font.Color = RGB(0, 128, 0)
            .Font.Bold = True
        End With
    End If
End Sub

' Validate cell formatting (whitespace and newlines) for all cells in a row
Private Sub ValidateCellFormatting(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, refValue As String, ByRef nextRow As Long)
    Dim lastCol As Long
    Dim colNum As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim trimmedText As String
    Dim colHeader As String
    Dim hasLeadingTrailingWS As Boolean
    Dim hasNewline As Boolean
    Dim hasDoubleSpace As Boolean
    Dim issuesList As String
    
    ' Get last column with data
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    
    ' Check each column in the row
    For colNum = 1 To lastCol
        cellValue = ws.Cells(rowNum, colNum).Value
        cellText = CStr(cellValue)
        
        ' Skip empty cells
        If cellText = "" Then GoTo NextColumn
        
        ' Check for leading/trailing whitespace
        trimmedText = Trim(cellText)
        hasLeadingTrailingWS = (Len(cellText) <> Len(trimmedText))
        
        ' Check for newlines (vbCrLf, vbLf, vbCr)
        hasNewline = (InStr(1, cellText, vbCrLf) > 0) Or _
                      (InStr(1, cellText, vbLf) > 0) Or _
                      (InStr(1, cellText, vbCr) > 0) Or _
                      (InStr(1, cellText, Chr(10)) > 0) Or _
                      (InStr(1, cellText, Chr(13)) > 0)
        
        ' Check for double spaces
        hasDoubleSpace = (InStr(1, cellText, "  ") > 0)
        
        ' Build issues list
        issuesList = ""
        If hasLeadingTrailingWS Then
            issuesList = "Leading/trailing whitespace"
        End If
        If hasNewline Then
            If issuesList <> "" Then issuesList = issuesList & ", "
            issuesList = issuesList & "Newline character(s)"
        End If
        If hasDoubleSpace Then
            If issuesList <> "" Then issuesList = issuesList & ", "
            issuesList = issuesList & "Double space(s)"
        End If
        
        ' Log issue if any formatting problems found
        If hasLeadingTrailingWS Or hasNewline Or hasDoubleSpace Then
            ' Get column header name
            colHeader = Trim(CStr(ws.Cells(1, colNum).Value))
            If colHeader = "" Then colHeader = "Column " & colNum
            
            LogIssue reviewWs, nextRow, "Formatting Issue", colHeader, refValue, cellText, issuesList
            nextRow = nextRow + 1
        End If
        
NextColumn:
    Next colNum
End Sub

' Validate blank cells within the used range for a row
Private Sub ValidateBlankCells(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, refValue As String, ByRef nextRow As Long)
    Dim lastCol As Long
    Dim colNum As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim colHeader As String
    
    ' Use the same column extent as the report (header row defines used columns)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For colNum = 1 To lastCol
        cellValue = ws.Cells(rowNum, colNum).Value
        cellText = Trim(CStr(cellValue))
        
        ' Treat as blank if empty or only whitespace
        If cellText = "" Then
            colHeader = Trim(CStr(ws.Cells(1, colNum).Value))
            If colHeader = "" Then colHeader = "Column " & colNum
            LogIssue reviewWs, nextRow, "Blank Cell", colHeader, refValue, "", "Cell is blank within used range"
            nextRow = nextRow + 1
        End If
    Next colNum
End Sub

' Get column index by header name (case-insensitive)
Private Function GetColumnIndexByHeader(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long
    Dim i As Long
    Dim cellValue As String
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        cellValue = Trim(CStr(ws.Cells(1, i).Value))
        If UCase(cellValue) = UCase(headerText) Then
            GetColumnIndexByHeader = i
            Exit Function
        End If
    Next i
    
    GetColumnIndexByHeader = 0 ' Not found
End Function

' Returns True only when term appears as a whole token (explicit match):
' case-sensitive, special characters as-is, bounded by non-word chars or start/end.
' e.g. "config" matches "the config is" but not "configuration".
Private Function TermAppearsAsExplicitMatch(text As String, term As String) As Boolean
    Dim termStr As String
    Dim p As Long
    Dim lenTerm As Long
    Dim lenText As Long
    Dim charBefore As String
    Dim charAfter As String
    
    termStr = CStr(term)
    If Len(termStr) = 0 Then
        TermAppearsAsExplicitMatch = False
        Exit Function
    End If
    
    lenTerm = Len(termStr)
    lenText = Len(text)
    p = 1
    
    Do
        p = InStr(p, text, termStr, vbBinaryCompare)
        If p = 0 Then Exit Do
        
        ' Character immediately before the match (or none)
        charBefore = ""
        If p > 1 Then charBefore = Mid(text, p - 1, 1)
        
        ' Character immediately after the match (or none)
        charAfter = ""
        If p + lenTerm <= lenText Then charAfter = Mid(text, p + lenTerm, 1)
        
        ' Term counts as explicit only if not surrounded by word characters
        If Not IsWordChar(charBefore) And Not IsWordChar(charAfter) Then
            TermAppearsAsExplicitMatch = True
            Exit Function
        End If
        
        p = p + 1
    Loop
    
    TermAppearsAsExplicitMatch = False
End Function

' True if single character is a word character (letter, digit, or underscore)
Private Function IsWordChar(c As String) As Boolean
    If Len(c) <> 1 Then
        IsWordChar = False
        Exit Function
    End If
    IsWordChar = (c Like "[A-Za-z0-9_]")
End Function

' Load forbidden terms from Review3 column A
Private Function LoadForbiddenTerms(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim terms() As String
    Dim i As Long
    Dim termCount As Long
    Dim cellValue As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        LoadForbiddenTerms = Array() ' Empty array
        Exit Function
    End If
    
    ReDim terms(1 To lastRow - 1) ' Rows 2 to lastRow
    termCount = 0
    
    For i = 2 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If cellValue <> "" Then
            termCount = termCount + 1
            terms(termCount) = cellValue
        End If
    Next i
    
    ' Resize array to actual count
    If termCount > 0 Then
        ReDim Preserve terms(1 To termCount)
        LoadForbiddenTerms = terms
    Else
        LoadForbiddenTerms = Array()
    End If
End Function

' Validate CG column (must be "Non Sub" or "RN")
Private Sub ValidateCGColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long)
    Dim cellValue As String
    
    cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    
    If cellValue <> "" Then
        If cellValue <> "Non Sub" And cellValue <> "RN" Then
            LogIssue reviewWs, nextRow, "Invalid CC", HEADER_CG, refValue, cellValue, HEADER_CG & " must be 'Non Sub' or 'RN'"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Validate Ref column (must start with CRQ followed by 6 or 7 digits)
Private Sub ValidateRefColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long)
    Dim cellValue As String
    Dim refPattern1 As String
    Dim refPattern2 As String
    Dim isValid As Boolean
    
    cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    
    If cellValue <> "" Then
        ' Check for CRQ followed by 6 digits
        refPattern1 = "CRQ######"
        ' Check for CRQ followed by 7 digits
        refPattern2 = "CRQ#######"
        
        isValid = (cellValue Like refPattern1) Or (cellValue Like refPattern2)
        
        If Not isValid Then
            LogIssue reviewWs, nextRow, "Invalid CRQ", HEADER_REF, refValue, cellValue, HEADER_REF & " must start with CRQ followed by 6 or 7 digits"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Validate Imp Date column (must be a valid date)
Private Sub ValidateImpDateColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long)
    Dim cellValue As Variant
    Dim cellText As String
    
    cellValue = ws.Cells(rowNum, colNum).Value
    cellText = Trim(CStr(cellValue))
    
    If cellText <> "" Then
        If Not IsDate(cellValue) Then
            LogIssue reviewWs, nextRow, "Invalid Imp Date", HEADER_IMP_DATE, refValue, cellText, HEADER_IMP_DATE & " must be a valid date (mm/dd/yyyy)"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Validate Desc column (spell check and forbidden terms)
Private Sub ValidateDescColumn(ws As Worksheet, reviewWs As Worksheet, keywordsWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long, forbiddenTerms As Variant)
    Dim cellValue As String
    Dim descText As String
    Dim words() As String
    Dim i As Long
    Dim misspelledWords As String
    Dim word As String
    Dim term As Variant
    Dim matchedTerms As String
    Dim hasSpellingIssue As Boolean
    Dim hasForbiddenTerm As Boolean
    
    cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    
    If cellValue = "" Then Exit Sub
    
    descText = cellValue
    misspelledWords = ""
    matchedTerms = ""
    
    ' Spell check: Split text into words and check each
    ' Note: Simple word extraction - split by spaces and common punctuation
    words = Split(Replace(Replace(Replace(Replace(descText, ",", " "), ".", " "), "(", " "), ")", " "), " ")
    
    For i = LBound(words) To UBound(words)
        word = Trim(words(i))
        If Len(word) > 0 Then
            ' Check spelling (Application.CheckSpelling returns False if word is misspelled)
            If Not Application.CheckSpelling(word) Then
                If misspelledWords <> "" Then misspelledWords = misspelledWords & ", "
                misspelledWords = misspelledWords & word
                hasSpellingIssue = True
            End If
        End If
    Next i
    
    ' Check for forbidden terms (explicit match only: whole token, case-sensitive, special chars as-is)
    If IsArray(forbiddenTerms) Then
        For Each term In forbiddenTerms
            If TermAppearsAsExplicitMatch(descText, CStr(term)) Then
                If matchedTerms <> "" Then matchedTerms = matchedTerms & ", "
                matchedTerms = matchedTerms & CStr(term)
                hasForbiddenTerm = True
            End If
        Next term
    End If
    
    ' Log spelling issues
    If hasSpellingIssue Then
        LogIssue reviewWs, nextRow, "Spelling", HEADER_DESC, refValue, misspelledWords, "Potential spelling issues found in " & HEADER_DESC
        nextRow = nextRow + 1
    End If
    
    ' Log forbidden term issues
    If hasForbiddenTerm Then
        LogIssue reviewWs, nextRow, "Forbidden Term", HEADER_DESC, refValue, matchedTerms, "Forbidden term(s) found in " & HEADER_DESC
        nextRow = nextRow + 1
    End If
End Sub

' Validate App Ref column based on CG value
Private Sub ValidateAppRefColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, cgCol As Long, appRefCol As Long, refValue As String, ByRef nextRow As Long)
    Dim cgValue As String
    Dim appRefValue As Variant
    Dim appRefText As String
    
    cgValue = Trim(CStr(ws.Cells(rowNum, cgCol).Value))
    appRefValue = ws.Cells(rowNum, appRefCol).Value
    appRefText = Trim(CStr(appRefValue))
    
    If cgValue = "Non Sub" Then
        ' App Ref must be "N/A"
        If appRefText <> "N/A" Then
            LogIssue reviewWs, nextRow, "CC/App Ref Mismatch", HEADER_APP_REF, refValue, appRefText, "When " & HEADER_CG & " is 'Non Sub', " & HEADER_APP_REF & " must be 'N/A'"
            nextRow = nextRow + 1
        End If
    ElseIf cgValue = "RN" Then
        ' App Ref must be a date
        If appRefText <> "" Then
            If Not IsDate(appRefValue) Then
                LogIssue reviewWs, nextRow, "CC/App Ref Mismatch", HEADER_APP_REF, refValue, appRefText, "When " & HEADER_CG & " is 'RN', " & HEADER_APP_REF & " must be a valid date"
                nextRow = nextRow + 1
            End If
        Else
            LogIssue reviewWs, nextRow, "CC/App Ref Mismatch", HEADER_APP_REF, refValue, "", "When " & HEADER_CG & " is 'RN', " & HEADER_APP_REF & " must contain a date"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Validate Outcome column (must be "Successful")
Private Sub ValidateOutcomeColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long)
    Dim cellValue As String
    
    cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    
    If cellValue <> "" Then
        If cellValue <> "Successful" Then
            LogIssue reviewWs, nextRow, "Invalid Results", HEADER_OUTCOME, refValue, cellValue, HEADER_OUTCOME & " must be 'Successful'"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Validate Conf column (must be "No")
Private Sub ValidateConfColumn(ws As Worksheet, reviewWs As Worksheet, rowNum As Long, colNum As Long, refValue As String, ByRef nextRow As Long)
    Dim cellValue As String
    
    cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
    
    If cellValue <> "" Then
        If cellValue <> "No" Then
            LogIssue reviewWs, nextRow, "Invalid Cert", HEADER_CONF, refValue, cellValue, HEADER_CONF & " must be 'No'"
            nextRow = nextRow + 1
        End If
    End If
End Sub

' Log an issue to the review table
Private Sub LogIssue(ws As Worksheet, rowNum As Long, issueType As String, colName As String, refValue As String, valueText As String, details As String)
    ws.Cells(rowNum, 1).Value = issueType
    ws.Cells(rowNum, 2).Value = colName
    ws.Cells(rowNum, 3).Value = refValue
    ws.Cells(rowNum, 4).Value = valueText
    ws.Cells(rowNum, 5).Value = details
End Sub

' Get the next available row for logging issues
Private Function GetNextIssueRow(ws As Worksheet) As Long
    GetNextIssueRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If GetNextIssueRow = 1 Then GetNextIssueRow = 2 ' Header row is row 1
End Function

' Format the review table
Private Sub FormatReviewTable(ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 1 Then Exit Sub
    
    Set tableRange = ws.Range("A1:E" & lastRow)
    
    ' Format header row with colorful styling
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) ' White text
        .Interior.Color = RGB(68, 114, 196) ' Blue header background
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Add borders to entire table
    With tableRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(100, 100, 100) ' Dark grey borders
    End With
    
    ' Add thicker border around entire table
    With tableRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With tableRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With tableRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With tableRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    
    ' Apply autofilter
    If lastRow > 1 Then
        tableRange.AutoFilter
    End If
    
    ' Auto-fit columns
    ws.Columns("A:E").AutoFit
    
    ' Apply alternating row colors with colorful styling (if there are issues)
    If lastRow > 1 Then
        Dim i As Long
        For i = 2 To lastRow
            If i Mod 2 = 0 Then
                ' Light blue/teal for even rows
                ws.Range("A" & i & ":E" & i).Interior.Color = RGB(230, 240, 255)
            Else
                ' White for odd rows
                ws.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 255, 255)
            End If
        Next i
    End If
End Sub


