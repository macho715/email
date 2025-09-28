VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Option Explicit

Private m_watchers As Collection

Private Sub Application_Startup()
    On Error GoTo CleanFail
    modHVDC_Config.InitializeHVDC
    Set m_watchers = New Collection

    Dim session As Outlook.NameSpace
    Set session = Application.Session

    Dim folders As Collection
    Set folders = modHVDC_Config.ResolveWatchedFolders(session)

    Dim folder As Outlook.MAPIFolder
    For Each folder In folders
        Dim watcher As clsHVDC_InboxWatcher
        Set watcher = New clsHVDC_InboxWatcher
        watcher.Initialise folder
        m_watchers.Add watcher
    Next folder

    Dim context As String
    context = "{""folder_count"":""" & Format$(folders.Count, "0.00") & """}"
    modHVDC_Logging.LogInfo "HVDC automation ready", context
    modHVDC_Config.RequestQueuePulse
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "Application_Startup", Err, ""
End Sub

Private Sub Application_Quit()
    On Error Resume Next
    Dim stats As HVDCBatchStats
    modHVDC_Export.ResetBatchStats stats
    Dim watcher As clsHVDC_InboxWatcher
    If Not m_watchers Is Nothing Then
        For Each watcher In m_watchers
            watcher.FlushPending stats
        Next watcher
    End If
    modHVDC_FileIO.PersistProcessedCache modHVDC_Config.HVDCProcessedCache
    Set m_watchers = Nothing
    modHVDC_Logging.LogInfo "HVDC automation stopped", "{}"
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    On Error GoTo CleanFail
    Dim entries As Variant
    entries = Split(EntryIDCollection, ",")
    Dim index As Long
    For index = LBound(entries) To UBound(entries)
        NotifyWatchers Trim$(CStr(entries(index)))
    Next index
    modHVDC_Config.RequestQueuePulse
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "Application_NewMailEx", Err, ""
End Sub

Private Sub Application_Reminder(ByVal Item As Object)
    On Error GoTo CleanFail
    Dim subjectText As String
    subjectText = ""
    If Not Item Is Nothing Then
        If TypeOf Item Is Outlook.TaskItem Then
            subjectText = Item.Subject
        Else
            Exit Sub
        End If
    End If

    If StrComp(subjectText, modHVDC_Config.QueueTaskName, vbTextCompare) = 0 Then
        HandleQueuePulse
    ElseIf StrComp(subjectText, modHVDC_Config.HourlyTaskName, vbTextCompare) = 0 Then
        HandleHourlyCatchUp
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "Application_Reminder", Err, ""
End Sub

Private Sub HandleQueuePulse()
    On Error GoTo CleanFail
    Dim stats As HVDCBatchStats
    modHVDC_Export.ResetBatchStats stats

    Dim watcher As clsHVDC_InboxWatcher
    Dim pending As Boolean
    pending = False
    If Not m_watchers Is Nothing Then
        For Each watcher In m_watchers
            watcher.ProcessPending stats, modHVDC_Config.BatchSizeLimit
            pending = pending Or watcher.HasPending
        Next watcher
    End If

    If stats.ProcessedCount + stats.SkippedCount + stats.ErrorCount > 0 Then
        modHVDC_Logging.LogBatchStats stats, "QueuePulse"
        modHVDC_FileIO.PersistProcessedCache modHVDC_Config.HVDCProcessedCache
    End If

    If pending Then
        modHVDC_Config.RequestQueuePulse
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "HandleQueuePulse", Err, ""
End Sub

Private Sub HandleHourlyCatchUp()
    On Error GoTo CleanFail
    modHVDC_Export.RunCatchUp
    modHVDC_Config.UpdateHourlyReminder
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "HandleHourlyCatchUp", Err, ""
End Sub

Private Sub NotifyWatchers(ByVal entryId As String)
    On Error GoTo CleanFail
    If Len(entryId) = 0 Then Exit Sub
    Dim watcher As clsHVDC_InboxWatcher
    If Not m_watchers Is Nothing Then
        For Each watcher In m_watchers
            watcher.NotifyEntry entryId
        Next watcher
    End If
    Exit Sub
CleanFail:
    modHVDC_Logging.LogException "NotifyWatchers", Err, entryId
End Sub

