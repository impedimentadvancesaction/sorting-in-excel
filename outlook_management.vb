Option Explicit

'==========================
' Configuration
'==========================

Private Function GetApprovedSenders() As Variant
    GetApprovedSenders = Array( _
        "firstname.lastname@example.com", _
        "firstname2.lastname2@example.com" _
    )
End Function

Private Function GetSavePath() As String
    GetSavePath = "C:\Your\HardCoded\Path\Here\"
End Function

Private Function GetTargetFolderPath() As String
    GetTargetFolderPath = "Inbox\TL-DM"
End Function

'==========================
' Entry point
'==========================

Public Sub RunProcessEmails()

    Dim inbox As Outlook.Folder
    Dim items As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim itm As Object

    Dim senderCounts As Object
    Set senderCounts = CreateObject("Scripting.Dictionary")

    InitialiseSenderCounts senderCounts

    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set items = inbox.Items

    items.Sort "[ReceivedTime]", True
    Set filteredItems = items.Restrict(GetYesterdayFilter())

    ' ---- EARLY EXIT ----
    If filteredItems.Count = 0 Then
        MsgBox "No emails found for yesterday from the configured senders.", _
               vbInformation, "Daily Email Processing"
        Exit Sub
    End If

    For Each itm In filteredItems
        If TypeOf itm Is Outlook.MailItem Then
            If IsApprovedSender(itm) Then
                ProcessMailItem itm
                IncrementSenderCount senderCounts, itm.SenderEmailAddress
            End If
        End If
    Next itm

    ShowSummary senderCounts

End Sub

'==========================
' Core logic
'==========================

Private Function GetYesterdayFilter() As String

    Dim startDate As Date
    Dim endDate As Date

    startDate = Date - 1
    endDate = Date

    GetYesterdayFilter = _
        "[ReceivedTime] >= '" & Format(startDate, "dd/mm/yyyy 00:00") & _
        "' AND [ReceivedTime] < '" & Format(endDate, "dd/mm/yyyy 00:00") & "'"

End Function

Private Function IsApprovedSender(mail As Outlook.MailItem) As Boolean

    Dim senders As Variant
    Dim sender As Variant

    senders = GetApprovedSenders()

    For Each sender In senders
        If LCase$(mail.SenderEmailAddress) = LCase$(CStr(sender)) Then
            IsApprovedSender = True
            Exit Function
        End If
    Next sender

    IsApprovedSender = False

End Function

Private Sub ProcessMailItem(mail As Outlook.MailItem)

    Dim savePath As String
    Dim fileName As String
    Dim targetFolder As Outlook.Folder

    savePath = GetSavePath()
    If Right$(savePath, 1) <> "\" Then savePath = savePath & "\"

    Set targetFolder = GetFolderFromPath(GetTargetFolderPath())

    fileName = BuildFileName(mail)

    mail.SaveAs savePath & fileName, olMSG
    mail.Move targetFolder

End Sub

Private Function BuildFileName(mail As Outlook.MailItem) As String

    Dim safeSubject As String

    safeSubject = mail.Subject
    safeSubject = Replace$(safeSubject, ":", "")
    safeSubject = Replace$(safeSubject, "\", "")
    safeSubject = Replace$(safeSubject, "/", "")
    safeSubject = Replace$(safeSubject, "*", "")
    safeSubject = Replace$(safeSubject, "?", "")
    safeSubject = Replace$(safeSubject, """", "")
    safeSubject = Replace$(safeSubject, "<", "")
    safeSubject = Replace$(safeSubject, ">", "")
    safeSubject = Replace$(safeSubject, "|", "")

    BuildFileName = _
        mail.SenderEmailAddress & " - " & _
        Trim$(safeSubject) & " - " & _
        Format$(mail.ReceivedTime, "dd-mm-yyyy") & ".msg"

End Function

'==========================
' Sender counting
'==========================

Private Sub InitialiseSenderCounts(dict As Object)

    Dim senders As Variant
    Dim sender As Variant

    senders = GetApprovedSenders()

    For Each sender In senders
        dict(LCase$(CStr(sender))) = 0
    Next sender

End Sub

Private Sub IncrementSenderCount(dict As Object, senderEmail As String)

    senderEmail = LCase$(senderEmail)

    If dict.Exists(senderEmail) Then
        dict(senderEmail) = dict(senderEmail) + 1
    End If

End Sub

Private Sub ShowSummary(dict As Object)

    Dim key As Variant
    Dim msg As String

    msg = "Emails processed:" & vbCrLf & vbCrLf

    For Each key In dict.Keys
        msg = msg & key & " : " & dict(key) & vbCrLf
    Next key

    MsgBox msg, vbInformation, "Daily Email Processing Summary"

End Sub

'==========================
' Folder resolution
'==========================

Private Function GetFolderFromPath(folderPath As String) As Outlook.Folder

    Dim parts As Variant
    Dim fldr As Outlook.Folder
    Dim i As Long

    parts = Split(folderPath, "\")

    Set fldr = Application.Session.GetDefaultFolder(olFolderInbox)

    For i = 1 To UBound(parts)
        Set fldr = fldr.Folders(CStr(parts(i)))
    Next i

    Set GetFolderFromPath = fldr

End Function
