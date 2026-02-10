Sub CleanReferenceList()
    Dim cell As Range
    For Each cell In Sheets("Sheet2").Range("A2:A100")
        If Not IsEmpty(cell) Then
            cell.Value = Replace(cell.Value, ChrW(8211), "-")
            cell.Value = Replace(cell.Value, ChrW(8212), "-")
            cell.Value = Replace(cell.Value, ChrW(8209), "-")
            cell.Value = Replace(cell.Value, ChrW(173), "-")
            cell.Value = Trim(cell.Value)
        End If
    Next cell
End Sub
