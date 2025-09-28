VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHVDC_InboxWatcher"
Option Explicit

Private WithEvents m_items As Outlook.Items
Attribute m_items.VB_VarHelpID = -1
Private m_queue As Collection
Private m_folder As Outlook.MAPIFolder
Private m_folderPath As String
Private m_storeId As String
Private m_pendingIndex As Object

Public Sub Initialise(ByVal targetFolder As Outlook.MAPIFolder)
    On Error GoTo CleanFail
    Set m_folder = targetFolder
    Set m_items = targetFolder.Items
    Set m_queue = New Collection
    m_folderPath = targetFolder.FolderPath
    m_storeId = targetFolder.StoreID
    Set m_pendingIndex = CreateObject("Scripting.Dictionary")
    m_pendingIndex.CompareMode = vbTextCompare
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.Initialise", Err, targetFolder.Name
End Sub

Public Sub NotifyEntry(ByVal entryId As String)
    On Error GoTo CleanFail
    If Len(Trim$(entryId)) = 0 Then
        Exit Sub
    End If
    EnqueueEntry entryId
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.NotifyEntry", Err, m_folderPath
End Sub

Public Sub FlushPending(ByRef stats As HVDCBatchStats)
    On Error GoTo CleanFail
    Do While m_queue.Count > 0
        ProcessPending stats, modHVDC_Config.BatchSizeLimit
    Loop
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.FlushPending", Err, m_folderPath
End Sub

Public Function HasPending() As Boolean
    If m_queue Is Nothing Then
        HasPending = False
    Else
        HasPending = (m_queue.Count > 0)
    End If
End Function

Public Sub ProcessPending(ByRef stats As HVDCBatchStats, ByVal limit As Long)
    On Error GoTo CleanFail
    If m_queue Is Nothing Then Exit Sub
    Dim processed As Long
    processed = 0
    Do While m_queue.Count > 0 And processed < limit
        Dim entryId As String
        entryId = CStr(m_queue(1))
        m_queue.Remove 1
        ProcessEntry entryId, stats
        processed = processed + 1
        If processed Mod 5 = 0 Then
            DoEvents
        End If
    Loop
    If m_queue.Count > 0 Then
        modHVDC_Config.RequestQueuePulse
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.ProcessPending", Err, m_folderPath
End Sub

Private Sub ProcessEntry(ByVal entryId As String, ByRef stats As HVDCBatchStats)
    On Error GoTo CleanFail
    Dim session As Outlook.NameSpace
    Set session = Application.Session
    Dim mail As Outlook.MailItem
    If Len(m_storeId) > 0 Then
        Set mail = session.GetItemFromID(entryId, m_storeId)
    Else
        Set mail = session.GetItemFromID(entryId)
    End If
    If Not mail Is Nothing Then
        modHVDC_Export.ProcessMailItem mail, stats
        If Not m_pendingIndex Is Nothing Then
            If m_pendingIndex.Exists(entryId) Then
                m_pendingIndex.Remove entryId
            End If
        End If
        Set mail = Nothing
    Else
        stats.ProcessedCount = stats.ProcessedCount + 1
        stats.SkippedCount = stats.SkippedCount + 1
    End If
    If Not m_pendingIndex Is Nothing Then
        If m_pendingIndex.Exists(entryId) Then
            m_pendingIndex.Remove entryId
        End If
    End If
    Exit Sub
CleanFail:
    stats.ErrorCount = stats.ErrorCount + 1
    If Not m_pendingIndex Is Nothing Then
        If m_pendingIndex.Exists(entryId) Then
            m_pendingIndex.Remove entryId
        End If
    End If
    modHVDC_Logging.LogException "InboxWatcher.ProcessEntry", Err, entryId
End Sub

Private Sub EnqueueEntry(ByVal entryId As String)
    On Error GoTo CleanFail
    If m_queue Is Nothing Then
        Set m_queue = New Collection
    End If
    If m_pendingIndex Is Nothing Then
        Set m_pendingIndex = CreateObject("Scripting.Dictionary")
        m_pendingIndex.CompareMode = vbTextCompare
    End If
    If m_pendingIndex.Exists(entryId) Then
        Exit Sub
    End If
    m_queue.Add entryId
    m_pendingIndex.Add entryId, True
    modHVDC_Config.RequestQueuePulse
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.EnqueueEntry", Err, entryId
End Sub

Private Sub m_items_ItemAdd(ByVal Item As Object)
    On Error GoTo CleanFail
    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item
        EnqueueEntry mail.EntryID
        Set mail = Nothing
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "InboxWatcher.ItemAdd", Err, m_folderPath
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_items = Nothing
    Set m_folder = Nothing
    Set m_queue = Nothing
    Set m_pendingIndex = Nothing
End Sub

