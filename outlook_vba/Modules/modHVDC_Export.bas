Attribute VB_Name = "modHVDC_Export"
Option Explicit

Public Type HVDCBatchStats
    ProcessedCount As Double
    ExportedCount As Double
    SkippedCount As Double
    ErrorCount As Double
    LatencyTotalMs As Double
End Type

Public Sub ResetBatchStats(ByRef stats As HVDCBatchStats)
    stats.ProcessedCount = 0
    stats.ExportedCount = 0
    stats.SkippedCount = 0
    stats.ErrorCount = 0
    stats.LatencyTotalMs = 0
End Sub

Public Sub ProcessMailItem(ByVal mail As Outlook.MailItem, ByRef stats As HVDCBatchStats)
    On Error GoTo CleanFail
    Dim entryId As String
    entryId = mail.EntryID
    If Len(entryId) = 0 Then
        entryId = "missing-entry-" & Format$(Now, "yyyymmddhhnnss")
    End If

    stats.ProcessedCount = stats.ProcessedCount + 1

    Dim cache As Object
    Set cache = modHVDC_Config.HVDCProcessedCache
    If cache.Exists(entryId) Then
        stats.SkippedCount = stats.SkippedCount + 1
        Exit Sub
    End If

    Dim messagePath As String
    messagePath = modHVDC_FileIO.ComposeMessageFilePath(mail)
    Dim datedFolder As String
    Dim folderIndex As Long
    folderIndex = InStrRev(messagePath, "\")
    If folderIndex > 0 Then
        datedFolder = Left$(messagePath, folderIndex - 1)
    Else
        datedFolder = modHVDC_Config.HVDCExportRoot
    End If

    If Dir$(messagePath) <> "" Then
        cache.Add entryId, True
        stats.SkippedCount = stats.SkippedCount + 1
        Exit Sub
    End If

    If Not modHVDC_FileIO.SaveMailItem(mail, messagePath) Then
        stats.ErrorCount = stats.ErrorCount + 1
        Exit Sub
    End If

    Dim attachmentsJson As String
    attachmentsJson = SerializeAttachments(mail, entryId, datedFolder)

    Dim metadataPath As String
    metadataPath = modHVDC_FileIO.ComposeMetadataFilePath(messagePath)
    Dim metadataJson As String
    metadataJson = BuildMetadataJson(mail, messagePath, attachmentsJson)
    modHVDC_FileIO.WriteJsonMetadata metadataJson, metadataPath

    If Not cache.Exists(entryId) Then
        cache.Add entryId, True
    End If

    stats.ExportedCount = stats.ExportedCount + 1
    stats.LatencyTotalMs = stats.LatencyTotalMs + ComputeLatencyMs(mail.ReceivedTime)
    Exit Sub
CleanFail:
    stats.ErrorCount = stats.ErrorCount + 1
    modHVDC_Logging.LogException "ProcessMailItem", Err, "entry=" & entryId
End Sub

Public Sub RunCatchUp()
    On Error GoTo CleanFail
    Dim stats As HVDCBatchStats
    ResetBatchStats stats

    Dim session As Outlook.NameSpace
    Set session = Application.Session

    Dim folders As Collection
    Set folders = modHVDC_Config.ResolveWatchedFolders(session)

    Dim checkpoint As Date
    checkpoint = modHVDC_FileIO.ReadCheckpointTimestamp

    Dim folder As Outlook.MAPIFolder
    For Each folder In folders
        ProcessCatchUpFolder folder, checkpoint, stats
    Next folder

    modHVDC_FileIO.WriteCheckpointTimestamp Now
    modHVDC_FileIO.PersistProcessedCache modHVDC_Config.HVDCProcessedCache
    modHVDC_Logging.LogBatchStats stats, "CatchUp"
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "RunCatchUp", Err, ""
End Sub

Private Sub ProcessCatchUpFolder(ByVal folder As Outlook.MAPIFolder, ByVal sinceTime As Date, ByRef stats As HVDCBatchStats)
    On Error GoTo CleanFail
    Dim items As Outlook.Items
    Set items = folder.Items
    If items Is Nothing Then Exit Sub

    items.IncludeRecurrences = False
    items.Sort "[ReceivedTime]", True

    Dim filter As String
    filter = "[ReceivedTime] >= '" & Format$(sinceTime - (1 / 1440#), "yyyy-mm-dd hh:nn") & "'"
    Dim recent As Outlook.Items
    Set recent = items.Restrict(filter)
    If recent Is Nothing Then Exit Sub

    Dim index As Long
    For index = 1 To recent.Count
        If TypeOf recent(index) Is Outlook.MailItem Then
            Dim mail As Outlook.MailItem
            Set mail = recent(index)
            ProcessMailItem mail, stats
            Set mail = Nothing
        End If
        If (index Mod modHVDC_Config.BatchSizeLimit) = 0 Then
            DoEvents
        End If
    Next index
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "ProcessCatchUpFolder", Err, folder.FolderPath
End Sub

Private Function SerializeAttachments(ByVal mail As Outlook.MailItem, ByVal entryId As String, ByVal datedFolder As String) As String
    On Error GoTo CleanFail
    Dim attachments As Outlook.Attachments
    Set attachments = mail.Attachments
    If attachments Is Nothing Or attachments.Count = 0 Then
        SerializeAttachments = "[]"
        Exit Function
    End If

    Dim attachmentFolder As String
    attachmentFolder = modHVDC_FileIO.ComposeAttachmentsFolder(entryId, datedFolder)

    Dim builder As String
    builder = "["
    Dim index As Long
    For index = 1 To attachments.Count
        Dim attachment As Outlook.Attachment
        Set attachment = attachments(index)
        Dim savedPath As String
        savedPath = modHVDC_FileIO.SaveAttachment(attachment, attachmentFolder)
        If Len(savedPath) > 0 Then
            If Len(builder) > 1 Then
                builder = builder & ","
            End If
            builder = builder & "{""name"":""" & JsonEscape(attachment.FileName) & """,""path"":""" & JsonEscape(savedPath) & """,""size"":""" & Format$(attachment.Size, "0.00") & """}"
        End If
        Set attachment = Nothing
    Next index
    builder = builder & "]"
    SerializeAttachments = builder
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "SerializeAttachments", Err, entryId
    SerializeAttachments = "[]"
End Function
Private Function BuildMetadataJson(ByVal mail As Outlook.MailItem, ByVal messagePath As String, ByVal attachmentsJson As String) As String
    Const DQ As String = Chr$(34)
    Dim receivedLocal As String
    If mail.ReceivedTime = 0 Then
        receivedLocal = Format$(Now, "yyyy-mm-dd\THH:nn:ss")
    Else
        receivedLocal = Format$(mail.ReceivedTime, "yyyy-mm-dd\THH:nn:ss")
    End If
    Dim receivedUtc As String
    receivedUtc = receivedLocal

    Dim payload As String
    payload = "{" & DQ & "entry_id" & DQ & ":" & DQ & JsonEscape(mail.EntryID) & DQ
    payload = payload & "," & DQ & "subject" & DQ & ":" & DQ & JsonEscape(mail.Subject) & DQ
    payload = payload & "," & DQ & "sender_name" & DQ & ":" & DQ & JsonEscape(mail.SenderName) & DQ
    payload = payload & "," & DQ & "sender_address" & DQ & ":" & DQ & JsonEscape(mail.SenderEmailAddress) & DQ
    payload = payload & "," & DQ & "received_local" & DQ & ":" & DQ & receivedLocal & DQ
    payload = payload & "," & DQ & "received_utc" & DQ & ":" & DQ & receivedUtc & DQ
    payload = payload & "," & DQ & "message_path" & DQ & ":" & DQ & JsonEscape(messagePath) & DQ
    payload = payload & "," & DQ & "to" & DQ & ":" & SerializeRecipients(mail, olTo)
    payload = payload & "," & DQ & "cc" & DQ & ":" & SerializeRecipients(mail, olCC)
    payload = payload & "," & DQ & "bcc" & DQ & ":" & SerializeRecipients(mail, olBCC)
    payload = payload & "," & DQ & "attachments" & DQ & ":" & attachmentsJson & "}"
    BuildMetadataJson = payload
End Function

Private Function SerializeRecipients(ByVal mail As Outlook.MailItem, ByVal recipientType As Outlook.OlMailRecipientType) As String
    On Error GoTo CleanFail
    Dim recipients As Outlook.Recipients
    Set recipients = mail.Recipients
    If recipients Is Nothing Then
        SerializeRecipients = "[]"
        Exit Function
    End If
    Dim builder As String
    builder = "["
    Dim index As Long
    For index = 1 To recipients.Count
        Dim recipient As Outlook.Recipient
        Set recipient = recipients(index)
        If recipient.Type = recipientType Then
            If Len(builder) > 1 Then
                builder = builder & ","
            End If
            builder = builder & "{""name"":""" & JsonEscape(recipient.Name) & """,""address"":""" & JsonEscape(recipient.Address) & """}"
        End If
        Set recipient = Nothing
    Next index
    builder = builder & "]"
    SerializeRecipients = builder
    Exit Function
CleanFail:
    modHVDC_Logging.LogException "SerializeRecipients", Err, ""
    SerializeRecipients = "[]"
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim sanitized As String
    sanitized = Replace(value, "\", "\\")
    sanitized = Replace(sanitized, """", "\"")
    sanitized = Replace(sanitized, vbCr, "\r")
    sanitized = Replace(sanitized, vbLf, "\n")
    JsonEscape = sanitized
End Function

Private Function ComputeLatencyMs(ByVal receivedTime As Date) As Double
    If receivedTime = 0 Then
        ComputeLatencyMs = 0
    Else
        ComputeLatencyMs = (Now - receivedTime) * 86400000#
    End If
End Function
